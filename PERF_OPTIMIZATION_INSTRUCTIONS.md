# 1llum1n4t1s.AlphaFS パフォーマンス最適化指示書

## 概要

Lhamielプロジェクトでの利用を前提に、ファイルシステム操作のパフォーマンスボトルネックを分析した結果に基づく修正指示。AlphaFS は Win32 API を直接呼び出して長パス対応等を実現するライブラリ。

---

## A1: 🔴 最優先 — GetFullPathCore で毎回 StringBuilder(32700) 確保

### 問題

`GetFullPathCore` の初期バッファサイズが `NativeMethods.MaxPathUnicode`（32700）固定。大多数のパスは260文字以下にもかかわらず、毎呼び出しで約64KBのヒープを確保する。ファイル列挙時はファイルごとにこれが発生する。

### 対象ファイル

- `src/AlphaFS/Filesystem/Path Class/Path Core Methods/Path.GetFullPathCore.cs`

### 現在のコード

```csharp
uint bufferSize = NativeMethods.MaxPathUnicode;  // 32700

startGetFullPathName:
    var buffer = new StringBuilder((int) bufferSize);
    var returnLength = ...GetFullPathName(pathLp, bufferSize, buffer, IntPtr.Zero);
    if (returnLength != Win32Errors.NO_ERROR)
    {
        if (returnLength > bufferSize)
        {
            bufferSize = returnLength;
            goto startGetFullPathName;  // リトライ
        }
    }
```

### 修正方針

初期バッファを `MaxPath`（260）にして、不足時のみ拡大する。既存の `goto startGetFullPathName` リトライロジックがあるので安全:

```csharp
// 変更前
uint bufferSize = NativeMethods.MaxPathUnicode;

// 変更後
uint bufferSize = NativeMethods.MaxPath;  // 260 — 大多数のパスはこれで収まる
```

### 注意事項

- 既存の `goto startGetFullPathName` リトライロジックにより、260文字を超えるパスは `returnLength > bufferSize` で検出されて自動拡大される
- 長パスの場合は1回余分にP/Invokeが走るが、短パスの大多数（99%以上）でのアロケーション削減メリットの方が圧倒的に大きい
- `GetFullPathName` Win32 API は必要バッファサイズを返すため、2回目は正確なサイズで確保できる

---

## A2: 🔴 最優先 — FindFileSystemEntryInfo でフィルター前に FileSystemEntryInfo を生成

### 問題

`FindFileSystemEntryInfo` の列挙ループで、`InclusionFilter` などのフィルタリングを適用する前に `NewFilesystemEntry` で `FileSystemEntryInfo` オブジェクトを毎エントリ生成している。フィルターで大量のエントリを除外するケースでは無駄なオブジェクト生成が多発する。

### 対象ファイル

- `src/AlphaFS/Filesystem/FindFileSystemEntryInfo.cs` — `Enumerate<T>` メソッド内のループ

### 修正方針

フィルタリングロジックを `NewFilesystemEntry` の前に移動する。具体的には:

1. `_nameFilter`（Regex パターンマッチ）は `WIN32_FIND_DATA` のファイル名だけで判定可能なので、オブジェクト生成前にチェックする
2. `InclusionFilter` がカスタムフィルターの場合はオブジェクト生成が必要だが、`_nameFilter` のみの場合は事前チェックで大幅に削減できる

```csharp
// 列挙ループ内で、NewFilesystemEntry の前に _nameFilter チェックを追加
var fileName = win32FindData.cFileName;

// _nameFilter がある場合はファイル名だけで事前フィルタリング
if (_nameFilter != null && !_nameFilter.IsMatch(fileName))
    continue;  // オブジェクト生成をスキップ

var fsei = NewFilesystemEntry(pathLp, fileName, win32FindData);
// ... 以降の処理
```

### 注意事項

- `InclusionFilter` や `ErrorFilter` 等のカスタムフィルターは `FileSystemEntryInfo` のプロパティにアクセスするため、これらが設定されている場合はオブジェクト生成が必要
- `_nameFilter` はワイルドカードパターンから生成された Regex であり、ファイル名文字列だけで評価可能
- `IsDirectory` フラグと `DirectoryEnumerationOptions.Files`/`Folders` の照合もオブジェクト生成前に `WIN32_FIND_DATA.dwFileAttributes` から直接判定可能

---

## A3: 🟡 中優先 — NormalizePath で毎回 StringBuilder(260) 確保

### 問題

`NormalizePath` と `NormalizePathDotSpaceHandler` が毎呼び出しで `new StringBuilder(NativeMethods.MaxPath)` （260文字）を確保。`GetFullPathCore` 内で呼ばれるため、高頻度で実行される。

### 対象ファイル

- `src/AlphaFS/Filesystem/Path Class/Path.Helpers.cs` — `NormalizePath()` と `NormalizePathDotSpaceHandler()`

### 現在のコード

```csharp
private static string NormalizePath(string path, GetFullPathOptions options)
{
    var newBuffer = new StringBuilder(NativeMethods.MaxPath);  // 毎回 260文字分
    ...
    return newBuffer.ToString();
}

private static StringBuilder NormalizePathDotSpaceHandler(string path, ...)
{
    var newBuffer = new StringBuilder(NativeMethods.MaxPath);  // 最大2文字しか返さないのに260文字
    ...
}
```

### 修正方針

**NormalizePath**: バッファを入力パスの長さベースで確保する:

```csharp
// 変更前
var newBuffer = new StringBuilder(NativeMethods.MaxPath);

// 変更後
var newBuffer = new StringBuilder(path.Length);
```

**NormalizePathDotSpaceHandler**: 返す文字列は "." か ".." 程度なので小さなバッファで十分:

```csharp
// 変更前
var newBuffer = new StringBuilder(NativeMethods.MaxPath);

// 変更後
var newBuffer = new StringBuilder(4);  // "." or ".." + 余裕
```

### 注意事項

- `NormalizePath` の出力は入力パスと同じ長さ以下になるため、`path.Length` で十分
- `NormalizePathDotSpaceHandler` はカレントディレクトリプレフィックス（"." or ".."）のみを返す

---

## A4: 🟡 中優先 — GetRegularPathCore の重複 StartsWith チェック連鎖

### 問題

`GetRegularPathCore` で5つの `StartsWith` チェックを直列実行し、最後の分岐ではさらに `StartsWith` + `Substring` で新しい文字列を生成。ホットパスで連鎖して呼ばれる。

### 対象ファイル

- `src/AlphaFS/Filesystem/Path Class/Path Core Methods/Path.GetRegularPathCore.cs`

### 現在のコード

```csharp
internal static string GetRegularPathCore(string path, GetFullPathOptions options, bool allowEmpty)
{
    // null/empty チェック...
    // options 適用...

    if (path.StartsWith(DosDeviceUncPrefix, StringComparison.OrdinalIgnoreCase))   // 1回目
        return UncPrefix + path.Substring(DosDeviceUncPrefix.Length);

    if (path.StartsWith(LogicalDrivePrefix, StringComparison.Ordinal))             // 2回目
        return path.Substring(LogicalDrivePrefix.Length);

    if (path.StartsWith(NonInterpretedPathPrefix, StringComparison.Ordinal))       // 3回目
        return path.Substring(NonInterpretedPathPrefix.Length);

    return path.StartsWith(GlobalRootPrefix, ...) ||                               // 4回目
           path.StartsWith(VolumePrefix, ...) ||                                   // 5回目
           !path.StartsWith(LongPathPrefix, ...)                                   // 6回目
        ? path
        : (path.StartsWith(LongPathUncPrefix, ...) ? ...);                         // 7回目
}
```

### 修正方針

最も一般的なケースを先にチェックして早期リターンする。大多数のパスは `\\?\` プレフィックスを持たないため、LongPathPrefix チェックを最初に置く:

```csharp
internal static string GetRegularPathCore(string path, GetFullPathOptions options, bool allowEmpty)
{
    // null/empty チェック（変更なし）...
    if (options != GetFullPathOptions.None)
        path = ApplyFullPathOptions(path, options);

    // 最も一般的なケース: LongPathPrefix で始まらないパスは即座に返す
    if (!path.StartsWith(LongPathPrefix, StringComparison.Ordinal))
        return path;

    // 以下、LongPathPrefix で始まるパスのみ到達
    if (path.StartsWith(LongPathUncPrefix, StringComparison.OrdinalIgnoreCase))
        return UncPrefix + path.Substring(LongPathUncPrefix.Length);

    if (path.StartsWith(DosDeviceUncPrefix, StringComparison.OrdinalIgnoreCase))
        return UncPrefix + path.Substring(DosDeviceUncPrefix.Length);

    if (path.StartsWith(GlobalRootPrefix, StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith(VolumePrefix, StringComparison.OrdinalIgnoreCase))
        return path;

    return path.Substring(LongPathPrefix.Length);
}
```

### 注意事項

- `LongPathPrefix` は `\\?\` で、他のプレフィックス（`DosDeviceUncPrefix` = `\\?\UNC\`、`LongPathUncPrefix` = `\\?\UNC\` 等）はすべて `LongPathPrefix` で始まる。したがって `!StartsWith(LongPathPrefix)` の早期リターンで他のチェックを全てスキップできる
- `LogicalDrivePrefix` と `NonInterpretedPathPrefix` が `LongPathPrefix` で始まるかどうかを実際の定数値で確認すること。始まらない場合は別途チェックが必要
- 文字列結合 `UncPrefix + path.Substring(...)` は `string.Concat` に最適化されるが、`ReadOnlySpan` + `string.Create` でさらにアロケーション削減が可能（.NET Framework ターゲットでは不可、.NET Core 以降のみ）

---

## A5: 🟡 中優先 — GetLongPathCore の文字列結合

### 問題

`GetLongPathCore` でプレフィックス付与時に `string +` 演算子で新しい文字列を生成。

### 対象ファイル

- `src/AlphaFS/Filesystem/Path Class/Path Core Methods/Path.GetLongPathCore.cs`

### 現在のコード

```csharp
if (path.StartsWith(UncPrefix, StringComparison.Ordinal))
    return LongPathUncPrefix + path.Substring(UncPrefix.Length);  // 2回のアロケーション

return IsPathRooted(path, false) && IsLogicalDriveCore(path, false, PathFormat.LongFullPath)
    ? LongPathPrefix + path   // アロケーション
    : path;
```

### 修正方針

`string.Concat` で明示的に結合するか、.NET Core 以降であれば `string.Create` を使う:

```csharp
if (path.StartsWith(UncPrefix, StringComparison.Ordinal))
    return string.Concat(LongPathUncPrefix, path.AsSpan(UncPrefix.Length));

return IsPathRooted(path, false) && IsLogicalDriveCore(path, false, PathFormat.LongFullPath)
    ? string.Concat(LongPathPrefix, path)
    : path;
```

### 注意事項

- `path.AsSpan(...)` は .NET Standard 2.1 / .NET Core 2.1 以降で利用可能。ターゲットフレームワークを確認すること
- .NET Framework をターゲットとしている場合は `string.Concat(LongPathUncPrefix, path.Substring(UncPrefix.Length))` のままで十分（コンパイラ最適化で `Substring + Concat` より `Concat` 1回で済む場合がある）

---

## A6: 🟡 中優先 — CountFileSystemObjects の LINQ .Count() 全走査

### 問題

全オーバーロードで `EnumerateFileSystemEntryInfosCore<string>(...).Count()` を呼んでおり、カウントのためだけに全エントリを文字列として生成・走査する。

### 対象ファイル

- `src/AlphaFS/Filesystem/Directory Class/Directory.CountFileSystemObjects.cs`

### 現在のコード

```csharp
public static long CountFileSystemObjects(string path, DirectoryEnumerationOptions options)
{
    return EnumerateFileSystemEntryInfosCore<string>(null, null, path,
        Path.WildcardStarMatchAll, null, options, null, PathFormat.RelativePath).Count();
}
```

### 修正方針

カウント専用の内部メソッドを追加するか、ジェネリック型引数を最軽量のものに変更する:

**案1（推奨）: 専用カウントメソッド**

`FindFileSystemEntryInfo` に `CountEntries()` メソッドを追加し、`yield return` せずにカウンターだけ増分する。

```csharp
// FindFileSystemEntryInfo に追加
internal long CountEntries()
{
    long count = 0;
    // Enumerate と同じ WIN32_FIND_DATA ループだがオブジェクト生成なし
    // FindFirstFileEx / FindNextFile を回してフィルタ条件に合うものだけカウント
    ...
    return count;
}
```

**案2（簡易）: そのまま維持**

`Count()` は `IEnumerable<T>` の遅延評価を走査するだけなので、各アイテムの `string` 生成コスト分の無駄はあるが大幅な改修は不要。A2（フィルター前のオブジェクト生成回避）の修正で間接的に改善される。

### 注意事項

- 案1は `FindFileSystemEntryInfo` の列挙ロジックと重複コードが生まれるリスクがある
- 既に列挙ロジックが複雑なため、案2で他の最適化の効果を見てから判断するのが安全

---

## A7: 🟡 中優先 — Regex オブジェクトが毎 Enumerate で生成

### 問題

`FindFileSystemEntryInfo` の `SearchPattern` セッターで、パターンが変わるたびに `new Regex(...)` を生成。同じパターン（特に `*`）で何度も列挙する場合にコンパイルコストが無駄になる。

### 対象ファイル

- `src/AlphaFS/Filesystem/FindFileSystemEntryInfo.cs` — `SearchPattern` プロパティのセッター

### 修正方針

よく使われるパターンの Regex をキャッシュする:

```csharp
// static フィールドに追加
private static readonly ConcurrentDictionary<string, Regex> _regexCache = new();

// SearchPattern セッター内
_nameFilter = searchPatternIsAll ? null
    : _regexCache.GetOrAdd(_searchPattern, pattern =>
        new Regex(string.Format(..., Regex.Escape(pattern)...), RegexOptions.IgnoreCase | RegexOptions.Compiled));
```

### 注意事項

- `RegexOptions.Compiled` はJITコンパイルされるため初回コストは高いが、同じパターンで大量マッチングする場合は大幅に高速化
- キャッシュが無限に増加しないよう、キャッシュサイズ上限を設けるか、よく使われるパターン（`*`, `*.*`）のみ事前キャッシュする
- `ConcurrentDictionary` を使う場合は `System.Collections.Concurrent` の using を追加

---

## A8: 🟡 中優先 — DefaultFileBufferSize = 4096 が小さい

### 問題

`NativeMethods.DefaultFileBufferSize = 4096` は現代のストレージで小さすぎ、`FindFileSystemEntryInfo.Enumerate` の `Queue` 初期容量にも流用されている。

### 対象ファイル

- `src/AlphaFS/Filesystem/Native Methods/NativeMethods.Constants.cs`

### 修正方針

ファイルI/O用バッファサイズを別定数に分離し、適切な値にする:

```csharp
// 変更前
public const int DefaultFileBufferSize = 4096;

// 変更後
public const int DefaultFileBufferSize = 65536;  // 64KB — 現代のSSD/HDDに適したサイズ

// Queue 初期容量として使われている箇所は別定数に分離
internal const int DefaultDirectoryQueueCapacity = 64;  // ディレクトリ列挙キュー用
```

### 注意事項

- `DefaultFileBufferSize` が使われている全箇所を grep で確認し、Queue 初期容量として使われている箇所は `DefaultDirectoryQueueCapacity` に置き換える
- ファイルコピー操作でバッファサイズとして使われている箇所は 65536 で問題ない
- 必要に応じてメモリとのトレードオフを考慮（組込み環境等）

---

## 修正順序の推奨

1. **A1** — GetFullPathCore バッファ最適化（1行変更、最大効果）
2. **A3** — NormalizePath バッファ最適化（2行変更、A1と合わせて効果倍増）
3. **A2** — フィルター前の事前チェック（中程度の変更、列挙性能に大きく影響）
4. **A4** — GetRegularPathCore 早期リターン（中程度の変更）
5. **A5** — GetLongPathCore 文字列結合最適化（小変更）
6. **A8** — バッファサイズ定数の最適化（定数変更 + grep）
7. **A7** — Regex キャッシュ化（中程度の変更）
8. **A6** — CountFileSystemObjects（A2の効果を見てから判断）

## テスト

既存のテストプロジェクトで全テストが通ることを確認する:

```bash
dotnet test
```

特に以下を重点的に確認:
- 長パス（260文字超）のファイル操作が正常に動作すること（A1 のバッファ縮小で問題ないか）
- ワイルドカードパターンでの列挙が正しく動作すること（A2, A7）
- UNC パス（`\\server\share`）の変換が正しいこと（A4）
