Changelog
=========

## [3.0.2] - 2026-08-30

ファイル・ディレクトリのコピー／移動とリパースポイント処理、およびネイティブ資源の寿命管理に関する不具合を修正しました。

### 🐛 不具合修正

- `Directory.Copy` / `Move` がコピー元配下をコピー先に指定すると再帰的に自己コピーする問題を修正し、処理開始前に拒否するよう変更
- ジャンクションをリンクとしてコピーするとき、コピー元ルートと入れ子のリンクを一貫して保持し、リンク先が存在しない場合も正しく複製するよう修正
- クロスボリュームなどでエミュレートされる `Directory.Move` がコピーオプションと進捗通知を失う問題、および置換失敗時にコピー元を先に削除し得る問題を修正
- `CopyOptions.CopyTimestamps` を指定した再帰コピーで、入れ子のディレクトリにもタイムスタンプを反映するよう修正
- `PathFormat.LongFullPath` を使う `File.Copy` / `Move` がコピーオプション、タイムスタンプ、進捗通知を不正な引数として拒否する問題を修正
- `File.Move` の置換時に、読み取り専用または隠し属性の既存ファイルを上書きできない問題を修正
- seek不能なハンドルを Append モードで開いたときと、`BackupFileStream` の終了処理でネイティブハンドルが残る問題を修正
- SMB／DFS 情報が公開するセキュリティ記述子のメモリが列挙バッファ解放後に無効になる問題を修正

## [3.0.1] - 2026-08-23

再帰的なファイルシステム操作のリパースポイント安全性と、コピー・ACL・長パス処理の不具合を修正しました。

### 🐛 不具合修正

- 再帰的な `Directory.Delete` / `DeleteEmptySubdirectories` / `Encrypt` / `Decrypt` / `Compress` / `Decompress` が、ツリー内のジャンクションやディレクトリシンボリックリンクの参照先へ入らないよう修正
  - 削除・暗号化・圧縮の対象として指定したツリー外のファイルを削除・変更する危険を解消
- `Directory.Copy` が空ディレクトリまたはファイルだけのディレクトリをコピーすると、コピー先ルートを作成しない問題を修正
  - `CopyOptions.CopySymbolicLink` を通常ディレクトリへ指定した場合は通常のツリーコピーを続行し、入れ子のディレクトリシンボリックリンクだけをリンクとして複製
- `Directory.Compress` / `Decompress` が長い子パスへ通常パスを `LongFullPath` として渡し、260 文字超で失敗する問題を修正
- DACL のみを指定した `File.Create` が不要な `SeSecurityPrivilege` を要求する問題と、`BackupFileStream.GetAccessControl` が P/Invoke 返却バッファの容量検査で失敗する問題を修正
  - SACL を含む場合だけ `SeSecurityPrivilege` と `ACCESS_SYSTEM_SECURITY` を要求
- `Path.GetFullPath` のネイティブ処理が失敗した際、実際の Win32 エラーではなく成功コード 0 から `NotImplementedException` を生成する問題を修正

## [3.0.0] - 2026-07-27

ACL 設定 API の既定挙動を `System.IO` に合わせる破壊的変更のみのリリースです。

### 💥 破壊的変更

- `SetAccessControl` の既定セクションを `AccessControlSections.All` から `AccessControlSections.Access` (DACL のみ) へ変更
  - 対象: `Directory.SetAccessControl(path, security[, pathFormat])` / `File.SetAccessControl` /
    `DirectoryInfo.SetAccessControl(security)` / `FileInfo.SetAccessControl(security)` の
    `includeSections` を取らない全 8 オーバーロード
  - 対になる `GetAccessControl` は元から DACL・所有者・グループしか読んでおらず、
    読んでもいない SACL まで書きに行く非対称な実装だった
  - さらに所有者とグループの書き込みには対象に対する `WRITE_OWNER` が必要なため、
    `GetAccessControl` → ルール追加 → `SetAccessControl` という最も一般的な流れが、
    所有者が自分ではない環境 (昇格プロセスが作成したディレクトリの所有者は `BUILTIN\Administrators`) で
    `(5) Access is denied` で失敗していた。`System.IO` は変更されたセクションだけを書くためこの問題は起きず、
    ドロップイン代替として挙動を揃えた
  - 所有者・グループ・監査 (SACL) を書き込む場合は `includeSections` を取るオーバーロードを使う
  - `BackupFileStream.SetAccessControl` は `GetAccessControl` が全セクションを読むバックアップ用途のため `All` のまま

## [2.1.2] - 2026-07-27

テストの環境判定のみの変更です。ライブラリ本体 (`src/`) に変更はなく、利用者から見た挙動は 2.1.1 と同一です。

### 🧪 テスト

- 拒否 ACE の環境判定 (`RequireDenyAclRoundTrip`) を、被テスト対象ではなく `System.IO` の ACL API で行うよう変更
  - 従来は `Alphaleonis.Win32.Filesystem.Directory.GetAccessControl` / `SetAccessControl` でプローブし、
    例外を握り潰していたため、AlphaFS 側の回帰が「環境が非対応」に化けて該当 3 テストが黙って skip されていた
  - `System.IO` で判定することで、AlphaFS だけが失敗するケースは skip されずテスト失敗として表面化する
  - 対象: `Directory.Delete_DirectoryHasDenyPermission` /
    `Directory.Move_UserExplicitDenyOnDestinationFolder` /
    `AlphaFS_Directory.Copy_UserExplicitDenyOnDestinationFolder`
- 同じく skip 時のメッセージに、握り潰していた例外の型とメッセージを含めるよう変更
  - CI 上で skip されたときに、環境要因なのか実装の問題なのかを後から判断できるようにするため
- 上記の変更により CI で実際に走り始めた 3 テストが、DACL のみを更新するよう修正
  - 拒否 ACE の付与・解除に `Directory.SetAccessControl(path, security)` の既定オーバーロードを
    使っていたが、この既定は `AccessControlSections.All` を書き込む。読み取り側の
    `Directory.GetAccessControl(path)` は `Access | Group | Owner` しか読まないため、
    読んでいない SACL と変更していない Owner まで書きに行き、所有者が `BUILTIN\Administrators`
    になる昇格環境では `(5) Access is denied` で失敗していた
  - 意図どおり `AccessControlSections.Access` を明示するオーバーロードに変更
    (リポジトリ内の他の ACL テストは元からこの形)

## [2.1.1] - 2026-07-27

v2.1.0 のリリース報告で「未完了」として残していた項目への対応です。公開 API の変更はありません。

### 🐛 不具合修正

- netapi32 が確保したバッファを `NetApiBufferFree` で解放するようになりました
  - 従来は `LocalFree` (`Marshal.FreeHGlobal`) で解放しており、netapi32 の内部ヒープ実装に
    たまたま一致していただけで、Win32 の契約としては未定義動作だった
  - 専用の `SafeNetApiBufferHandle` を追加し、`NetShareEnum` / `NetSessionEnum` / `NetFileEnum` /
    `NetConnectionEnum` / `NetServerDiskEnum` / `NetDfsEnum` / `NetShareGetInfo` / `NetStatisticsGet` /
    `NetDfsGetInfo` / `NetDfsGetClientInfo` の 10 本すべてを移行
  - 呼び出し側が確保する `WNetGetUniversalName` (mpr.dll) のバッファは従来どおり

### 🧪 テスト

- **TxF (Transacted) API のスモークテストを追加** (12 件)
  - `KernelTransaction` を取る API はソース 109 ファイル・パラメータ 601 箇所に及ぶ一方、
    テストが 1 件も存在せず、通常版の変更が Transacted 版を壊しても検知できない状態だった
  - 対象: `File.WriteAllText` / `ReadAllText` / `GetSize` / `Copy` / `Move` / `Delete` / `Exists`、
    `Directory.CreateDirectory` (多階層含む) / `Exists` / `Delete` / `EnumerateFiles`
  - commit で反映されること、rollback および commit なしの Dispose で取り消されること、
    未コミットの変更がトランザクション外から見えないことを検証する
  - TxF が使えない環境 (ReFS / FAT32 等) は `UnitTestConstants.RequireTransactionalNtfs` で skip。
    判定には被テスト対象ではなく `System.IO.DriveInfo` を使う
- **ドライブ文字の取り合いによる不安定な失敗を修正** (2 件、根本原因は同一)
  - `DriveConnection` は「最後の空きドライブ文字」(通常 `Z:`) を割り当てるが、これを使う
    4 つのテストに `[DoNotParallelize]` が付いておらず、互いに、またドライブを列挙する
    テストと並列実行されていた
    - `AlphaFS_Host_DriveConnection_Network_Success` が `ERROR_ALREADY_ASSIGNED`
      (ローカルデバイス名は既に使用されています) で落ちる
    - `AlphaFS_Volume_GetVolumeLabel_Local_Success` が `DriveNotFoundException` で落ちる
      (列挙中に `Z:` が消える)
  - `AlphaFS_Host.DriveConnection` / `AlphaFS_Host.GetMappedConnectionName` /
    `AlphaFS_Host.GetMappedUncName` / `AlphaFS_Directory.CreateJunction_FromMappedDrive` に
    `[DoNotParallelize]` を追加し、ドライブ文字を割り当てる全 8 テストで指定が揃った
  - あわせて `AlphaFS_Volume_GetVolumeLabel_Local_Success` 自体の TOCTOU も解消
    - ドライブの列挙・`IsReady` の確認・`VolumeLabel` の読み取りが別々の時点で行われており、
      `VolumeLabel` は 2 回読んでいた。読み取りを 1 回にまとめ、System.IO 側の probe だけを
      ガードして、被テスト対象である `Volume.GetVolumeLabel` の失敗は握り潰さないようにした
      (ネットワークドライブの切断など、テスト外の要因にも耐えるようにするため)
- `Directory_Delete_ExistingDirectory_LocalAndNetwork_Success` に削除後の非存在検証を追加

### 🔒 CI

- `actions/checkout` と `actions/setup-dotnet` をフルコミット SHA へ固定
  - このジョブは `secrets.NUGET_API_KEY` を扱うため、書き換え可能なタグ参照を避ける
- テスト実行のコマンドに、引数を追加してはいけない理由をコメントとして明記
  - MTP はサポートしていない引数を渡されると、テストを 1 件も実行しないまま
    "Zero tests ran" / exit 5 で抜ける
- `.github/dependabot.yml` を追加 (月 1 回、GitHub Actions と NuGet)
  - SHA 固定した Action は手動では追随できないため、Dependabot に SHA とバージョンコメントを更新させる
  - minor / patch はグループ化して PR 数を抑え、メジャー更新だけ個別 PR にする
  - `MSTest.TestAdapter` と `MSTest.TestFramework` は常に同一バージョンへ揃うよう専用グループにする
  - NuGet はソリューション (`AlphaFS.slnx`) ではなくプロジェクトのディレクトリを直接指定する

## [2.1.0] - 2026-07-26

8 観点 + 検証 2 段のコードレビュー (/rere) で確定した指摘をまとめて修正しました。
Arm64 での利用不能、`BackupFileStream` のハンドルリークと代替データストリーム名の破損、
NuGet への誤バージョン公開を許す CI の穴が主な内容です。

### 🎉 新機能・対応環境

- **Arm64 ネイティブの .NET プロセスから利用できるようになりました**
  - `AlphaFS.csproj` の `<PlatformTarget>x64</PlatformTarget>` を削除 (AnyCPU へ)
  - x64 を刻んだアセンブリは PE のマシン種別が AMD64 で固定され、Arm64 プロセスでは
    `FileLoadException: The assembly architecture is not compatible with the current process architecture.`
    でロードできなかった。コード側に x64 固有の依存は無く、ポインタ幅依存の箇所は Arm64 でも正しく動作する
- `OperatingSystem.EnumProcessorArchitecture` に `Arm` (5) / `Arm64` (12) を追加
  - Arm64 環境で `ProcessorArchitecture` が未定義値 12 を返し、`Unknown` にすら落ちなかったのを是正 (既存の値は不変)
- NuGet パッケージに SourceLink とシンボルパッケージ (`.snupkg`) を同梱
  - P/Invoke ラッパーという性質上、利用者の障害スタックは AlphaFS 内部で止まるため、
    ファイル名・行番号が出せないと原因に到達できなかった

### 🐛 不具合修正

- `BackupFileStream.Dispose()` がファイルハンドルを閉じない問題を修正
  - `CloseSafeHandle` が `_context != IntPtr.Zero` の分岐内にあり、`BackupRead` / `BackupWrite` / `BackupSeek` を
    一度も呼んでいないインスタンスではハンドルが残っていた
  - パス指定コンストラクタの既定は `FileShare.None` なので、GC まで対象ファイルがロックされ続けていた
- `BackupFileStream.ReadStreamInfo()` が代替データストリーム名を 10 文字で打ち切る問題を修正
  - ヘッダー用の 20 バイトバッファを名前の読み取りに使い回していた
  - 読み残しが次の `WIN32_STREAM_ID` として誤解釈され、以降のヘッダー解析がすべて破綻していた
  - ダウンロードしたファイルに自動付与される `:Zone.Identifier:$DATA` (44 バイト) で確実に発生
- `BackupFileStream.Read()` が実際の読み取り長ではなく要求量を丸ごとコピーしていた問題を修正
  - 未初期化のネイティブヒープ内容が呼び出し元のバッファへ転写されていた
- `Host` の SMB / DFS 列挙が `ERROR_MORE_DATA` で無限ループする問題を修正
  - 再開ハンドルがデリゲート定義上 `out` のみで、次の反復へ渡す口が存在しなかった
- `Directory.CreateJunction` / `File.CreateSymbolicLink` が投げる `IOException` の `HResult` を HRESULT へ変換
  - 生の `183` は severity bit が 0 で HRESULT 規約上「成功」を意味していた
- `File.Copy` が既存ファイルに対して投げる `AlreadyExistsException` のメッセージが `(80)` ではなく `(183)` を表示していた問題を修正
- `File.CreateFileStreamCore` が、`FileStream` 構築の失敗時に有効なファイルハンドルを閉じずに伝播していた問題を修正
- `PrivilegeEnabler` が、元から有効だった特権をスコープ離脱時に無効化していた問題を修正
  - 巻き戻し判定が `TOKEN_PRIVILEGES.PrivilegeCount` だけを見ており、`Attributes` の `SE_PRIVILEGE_ENABLED` を検査していなかった
  - 公開ドキュメント「既に有効な特権は無効にされません。」の記述どおりの挙動になった
- `PrivilegeEnabler` の `UnauthorizedAccessException` に特権名が出ない問題を修正 (フィールドを先にクリアしていた)
- `NativeMethods.SYSTEM_INFO` のマーシャリング幅を修正
  - `wProcessorArchitecture` が enum の基底型 int (4 バイト) として扱われ、`dwPageSize` 以降がネイティブ配置とずれていた
- COM ラッパー 4 種の `Dispose` を `Interlocked.Exchange` 化し、並行 Dispose での二重 `Release` を防止
- `OperatingSystem` の初期化で、センチネルより後に代入されるフィールドがあったのを是正
  - 初回アクセスが競合すると `VersionName` が `Later`、`IsServer` が `false` を返し得た

### ⚠️ 挙動変更

- `File.IsLocked` / `FileInfo.IsLocked` が、パス検証由来の `ArgumentException` をそのまま伝播するようになりました
  - 従来は例外の HRESULT 下位 16bit を Win32 エラーコードとして再解釈し、
    `IOException: (87) パラメーターが間違っています` へ変換していた
  - `catch (ArgumentException)` を書いていたコードが正しく動くようになる一方、
    `catch (IOException)` だけで受けていたコードは例外が抜けるようになる
- `Alphaleonis.Win32.DriveInfo` の各プロパティが取得に失敗した際、`Trace` に警告を出力するようになりました
  - 戻り値 (0 / null / 空文字) は従来どおりで、公開契約は変更していない

### ⚡ 性能

- `Directory.CreateDirectory` が、既存の祖先ディレクトリを見つけた時点で走査を打ち切るようになりました
  - 従来はセグメント数 D に対し常に D 回のカーネル呼び出しと O(D×n) の文字列コピーを行っていた
  - ディレクトリツリーのコピーでは宛先フォルダ 1 個ごとにこの走査が発生していた
- `Directory.Copy` / `Move` が、ツリー内のファイルごとに宛先の親ディレクトリを作り直さなくなりました
  - ファイル N 個あたり `GetFullPathNameW` × N + `GetFileAttributesExW` × 2N を削減 (再試行時は従来どおり作り直す)
- `Path.IsLogicalDrive` がパス全体を大文字化しなくなりました
  - ほぼ全ての公開 API が通るホットパスで、拡張長パスでは 1 呼び出しあたり数十 KB を確保していた
- `Directory.DeleteEmptySubdirectories` が、確定済みの長絶対パスを再正規化しなくなりました
- `Host` のネットワーク列挙が `INetworkListManager` を都度生成・解放するようになりました
  - 一過性の COM 生成失敗が `TypeInitializationException` として永続化し、
    以降そのプロセスで列挙が使えなくなる問題を解消

### 🔒 CI / リリース

- `publish.yml` のバージョン照合を、`push` / `workflow_dispatch` のどちらの経路でも必ず実行するよう修正
  - 従来は `workflow_dispatch` で照合ステップが skip され (skip は success 扱い)、
    `release/**` 以外のブランチから未公開バージョンを nuget.org へ確定公開できた
  - NuGet はバージョン番号を永久に予約し、unlist しても同一バージョンの再アップロードはできないため不可逆だった
  - 照合を checkout 直後へ移動し、失敗を数分ではなく数秒で検出
- `publish.ps1` が `Directory.Build.props` の `<Version>` から push 対象を決め打ちするよう修正
  - 更新時刻が最新の `*.nupkg` を選ぶ方式では、`artifacts/` に別バージョンや別 ID の
    パッケージが残っていた場合にそれを公開し得た
  - `param([string]$ApiKey = $env:NUGET_API_KEY)` を追加し、PowerShell の動的スコープ依存を解消
  - API キーの文字数をログ出力するのをやめた

### 🧹 整理

- `Path.NormalizePath` 内の到達不能な UNC 検証ブロックを削除 (`const int result = 1` により実行時に何もしていなかった)
- 単一 TFM では常に一方しか選ばれない `#if NET35` を src / tests から除去 (16 箇所)
- 復元に使われていない `src/AlphaFS/packages.config` と、`<startup/>` のみの `src/AlphaFS/app.config` を削除
- `CLAUDE.md` の名前空間 → ディレクトリ対応表の誤りを修正 (`src/AlphaFS/Device/` は `Alphaleonis.Win32.Filesystem`)
- `CLAUDE.md` にテスト用環境変数 (`ALPHAFS_SKIP_NETWORK_TESTS` / `ALPHAFS_ENABLE_MACHINE_STATE_TESTS`) と、
  昇格環境では `RequireElevation` のテストが skip されない点を追記
- `README.md` に、現行 `System.IO` と食い違う API (`CreateSymbolicLink` の戻り値型、`LinkTarget` / `ResolveLinkTarget` の不在) の一覧を追記
- `OperatingSystem.EnumOsName` の順序保証に関するドキュメントを実態に合わせて訂正
  - クライアント系とサーバー系が交互に配置されているため、ライン跨ぎの序数比較は build 番号順と一致しない

### 🧪 テスト

- ランダム化テストヘルパーの欠陥を 2 件修正
  - `Random.Next(1, 3)` の上限が排他のため `case 3` (â/ê/î 系のパス名) が一度も選ばれていなかった
  - 属性設定で `new Random(DateTime.UtcNow.Millisecond)` を 2 つ作っていたため、同一ミリ秒内では
    ReadOnly と Hidden の判定が必ず一致し、「片方だけ」の組み合わせが生成されていなかった
- `TemporaryDirectory` の no-op なファイナライザを削除

## [2.0.0] - 2026-07-26

テストが実行されていなかった問題を修正したところ、プロセスを即死させる不具合を含む複数の潜在バグが表面化したため、まとめて修正しました。
あわせて .NET Framework 時代の挙動を引きずっていた箇所を現行の `System.IO` に合わせています (破壊的変更)。

### 💥 破壊的変更

- `Path.Combine("C:", "file")` の結果を `C:file` (ドライブ相対) から `C:\file` へ変更
- `DirectoryInfo.Parent` / `Root` / 列挙結果の `ToString()` が、名前のみではなくフルパスを返すよう変更
- `FileInfo` / `DirectoryInfo` の `Create()` / `Delete()` がキャッシュを自動で無効化するよう変更
  - `Refresh()` を呼ばなくても `Exists` が最新の値を返す
- `OperatingSystem.EnumOsName` に `WindowsServer2019` / `Windows11` / `WindowsServer2022` / `WindowsServer2025` を追加
  - Windows 11 以降を `Later` と判定していたのを是正。`Later` は「未知の将来 OS」専用の値になった (既存の値は不変)

`Path.GetFullPath` / `GetPathRoot` の拡張長パス (`\\?\`) 処理と、不正なパスに対する例外送出は、
AlphaFS が意図的に `System.IO` より厳格なまま維持しています。

### 🐛 不具合修正

- `Shell32Info` のプロパティ参照でプロセスがアクセス違反 (0xC0000005) で即死する問題を修正
  - `IQueryAssociations::GetString` を vtable インデックス 5 (実際は `GetKey`) で呼び出していた
  - `Init` に失敗した COM オブジェクトに対しても `GetString` を発行していた (拡張子の無いファイルなどで発生)
- `Process.GetCurrentProcess().Handle` をインスタンス未保持で使用していた 3 箇所を修正
  - GC による回収でハンドルが閉じられ、`ERROR_INVALID_HANDLE` になり得た
  - 対象: `OperatingSystem.IsWow64Process` / `ProcessContext` / `Path.GetFinalPathNameByHandleCore`

### 🧪 テスト

- MSTest 4.x / Microsoft.Testing.Platform (MTP) へ移行し、テストが 1 件も実行されていなかった状態を解消
  - テストプロジェクトに `EnableMSTestRunner` を設定し、リポジトリ直下に `global.json` を追加
  - CI のテスト実行から VSTest 専用の `--logger` を削除
- `Assert.IsGreaterThan` / `IsLessThan` の引数逆転 17 箇所を修正
- ネットワーク共有・管理者権限が利用できない環境では失敗ではなく skip するよう整備
  - 環境変数 `ALPHAFS_SKIP_NETWORK_TESTS=1` でネットワーク側の検証を明示的に無効化できる
- ドライブ文字を割り当てるテストを並列実行対象から除外し、テスト間干渉を解消

## [1.0.38] - 2026-07-26

### 📦 依存パッケージの更新
- System.Security.Permissions 10.0.8 → 10.0.10
- MSTest.TestAdapter / MSTest.TestFramework 4.2.3 → 4.3.2
- 脆弱性・非推奨パッケージはなし

### 🔧 ビルド構成
- バージョン定義をリポジトリ直下の `Directory.Build.props` に一元化
  - `AlphaFS.csproj` の `<Version>` を削除し、継承へ変更
  - CI のバージョン照合も `Directory.Build.props` を参照するよう更新

## [1.0.37] - 2026-05-17

### コードレビュー指摘の P0/P1/P3 修正 (/rere 12 人分隊レビュー)

#### 🛡️ セキュリティ・データ整合性 (Critical)
- COM ラッパー 7 クラスに finalizer + Dispose(bool) パターン追加 (NetworkWrapper / NetworkConnectionWrapper / NetworkListManagerWrapper / QueryAssociationsWrapper / NetworkInfo / NetworkConnectionInfo / Shell32Info)
  - NetworkInfo / NetworkConnectionInfo は sealed 化
  - AOT 環境で Dispose 漏れ時の COM 参照永久リーク確定の脆弱性を解消
- `AdjustTokenPrivileges` の `ERROR_NOT_ALL_ASSIGNED` 検出を追加
  - 非管理者ユーザで Backup/Restore 等を要求した場合のサイレント成功を防ぐ
- `PrivilegeEnabler` の部分構築失敗時の特権リーク防止 (try/catch で巻き戻し)
- `PrivilegeEnabler.Dispose` の catch にログ追加 (特権無効化失敗を可視化)

#### 🪟 OS 判定 / AOT 互換性
- Windows 11 / Server 2022 以降を Later 判定 (`dwBuildNumber` チェック追加)
  - 誤って `EnumOsName.Windows10` と返すバグを修正
- `Utils.GetEnumDescription` / `Device.EnumerateDevices` に `[RequiresUnreferencedCode]` 注釈
- `Win32Errors.ERROR_NOT_ALL_ASSIGNED` 定数を有効化

#### ⚡ パフォーマンス
- CRC32 / CRC64 ホットループの `IList<byte>` → `byte[]` 化 (仮想インタフェース呼び出し除去)

#### 📝 ドキュメント・設定整合性
- CLAUDE.md の完全二重重複を削除
- CLAUDE.md の `DefaultFileBufferSize = 4096` → `65536` 修正 (実装と整合)
- CLAUDE.md の Breaking Changes セクションに README へのリンク追加
- `*.csproj` の `<Authors>` セパレータを統一 (`;` と `,` 混在 → カンマ統一)
- 30 ファイル / 132 行の UTF-8 文字化け修正 (フォーク作業中の混入)
- `KernelTransaction` XML doc 内の COM 関連文字化けを修正
- MSTest.TestFramework / TestAdapter 4.1.0 → 4.2.3 へ更新
- `File.CreateTextCore` の `<returns>` XML doc 誤記を修正
- `NativeError.ThrowException` default ブランチに HResult 利用方針コメント追加

#### 🚀 CI/CD
- `publish.yml` に `dotnet test` ステップ追加
- `publish.yml` に `permissions: contents:read` 追加 (最小権限原則)
- `publish.yml` に csproj `<Version>` とブランチ名の整合チェック追加
- `publish.ps1` に `--skip-duplicate` 追加 (重複 push を CI 失敗と区別)
- csproj から `CleanBinObjBeforeBuild` Target を削除 (インクリメンタルビルド復活)
- `GenerateAppIcon` Target に Inputs/Outputs 追加 (アイコンキャッシュ有効化)

#### 🩺 運用・保守性
- `Volume.IsSameVolume` の `catch{}` に `Trace.TraceWarning` を追加 (`Directory.Move` の Copy+Delete fallback 誤発動の診断)
- `NativeMethods.Utilities.CloseSafeHandle` の死コード `handle = null` を削除
- `tests/AlphaFS.UnitTest/AlphaFS.UnitTest.csproj` の `<Version>1.0.16</Version>` を削除 (`IsPackable=false` のため不要)

## [1.0.36] — Git 記録日: 2026-04-05

- ハッシュ計算・正規表現・文字列処理の不要なメモリ割り当てを減らし、Native AOT との互換性を改善。既存の非推奨公開 API は互換性のため維持。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/03fd40a7fda186f8a18c871f446fcf489f3d31a6) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/95137642db2b0f1258f1251acafdb3a5daa6c963...03fd40a7fda186f8a18c871f446fcf489f3d31a6)。

## [1.0.34] — Git 記録日: 2026-04-05

- 正規表現キャッシュに上限を設け、メモリ使用量が増え続ける問題を修正。
- ファイル I/O とネイティブ問い合わせのバッファを分離し、不要なメモリ使用を削減。API の説明を日本語化。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/95137642db2b0f1258f1251acafdb3a5daa6c963) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/9f70bb45a90ef02f2dbc9293e1a768076fce82e5...95137642db2b0f1258f1251acafdb3a5daa6c963)。

## [1.0.32] — Git 記録日: 2026-03-10

- ディレクトリの日時設定がファイルとして処理される不具合を修正。
- パス処理・サイズ集計・削除処理のメモリ使用量を減らし、正規表現の Native AOT 互換性を改善。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/9f70bb45a90ef02f2dbc9293e1a768076fce82e5) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/d619a267b0e9e3446565a880fe14b01f1d47e80d...9f70bb45a90ef02f2dbc9293e1a768076fce82e5)。

## [1.0.30] — Git 記録日: 2026-02-10

- AOT対応

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/d619a267b0e9e3446565a880fe14b01f1d47e80d) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/04b8fe2b1ac1345c2f437bf6e0d03d7b72da5179...d619a267b0e9e3446565a880fe14b01f1d47e80d)。

## [1.0.28] — Git 記録日: 2026-02-09

- ネイティブAOT対応

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/04b8fe2b1ac1345c2f437bf6e0d03d7b72da5179) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/1c5ae0568d6c830a9476a50693b9a3447b4aea46...04b8fe2b1ac1345c2f437bf6e0d03d7b72da5179)。

## [1.0.20] — Git 記録日: 2026-02-07

- UnitSizeToText で極端に大きい値を渡した際の IndexOutOfRangeException を修正
- IsWow64Process のエラーハンドリングと Crc32 の演算子優先度を修正

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/1c5ae0568d6c830a9476a50693b9a3447b4aea46) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/0d774561fabddafaae09ec5c7de26002aec932d2...1c5ae0568d6c830a9476a50693b9a3447b4aea46)。

## [1.0.18] — Git 記録日: 2026-02-03

- 末尾のディレクトリ区切り文字の判定で、両方の区切り文字を認識するよう修正。内部処理とファイル構成を整理。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/0d774561fabddafaae09ec5c7de26002aec932d2) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/d3fb492fd0a72fc0185463750f3043ce85311af3...0d774561fabddafaae09ec5c7de26002aec932d2)。

## [1.0.16] — Git 記録日: 2026-02-01

- 相対パス計算のループ処理を修正し、頻繁に実行する処理での不要な割り当てを削減。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/d3fb492fd0a72fc0185463750f3043ce85311af3) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/2b31561dace567b8abeec529a09d19380592fb07...d3fb492fd0a72fc0185463750f3043ce85311af3)。

## [1.0.14] — Git 記録日: 2026-01-29

- アイコンとビルド設定を調整し、復元前に古いビルド成果物を消去する処理を追加。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/2b31561dace567b8abeec529a09d19380592fb07) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/d39c24f7fd3daf6cc16e8cc26dc865c573f3472a...2b31561dace567b8abeec529a09d19380592fb07)。

## [1.0.12] — Git 記録日: 2026-01-26

- アイコン設定

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/d39c24f7fd3daf6cc16e8cc26dc865c573f3472a) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/84fc32082eed68b660d56a063ba41f7be86387bb...d39c24f7fd3daf6cc16e8cc26dc865c573f3472a)。

## [1.0.10] — Git 記録日: 2026-01-25

- アイコン設定

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/84fc32082eed68b660d56a063ba41f7be86387bb) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/3cf5ddaec70cb9aa8360d5ca1935069c3f585cfd...84fc32082eed68b660d56a063ba41f7be86387bb)。

## [1.0.8] — Git 記録日: 2026-01-25

- x64固定

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/3cf5ddaec70cb9aa8360d5ca1935069c3f585cfd) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/aab32537460e07fe3bfa57cafd87de5de3350e2b...3cf5ddaec70cb9aa8360d5ca1935069c3f585cfd)。

## [1.0.5] — Git 記録日: 2026-01-25

- アイコン設定

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/aab32537460e07fe3bfa57cafd87de5de3350e2b) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/b8430cbcc6b50f89ccac25d91167819b27c05c22...aab32537460e07fe3bfa57cafd87de5de3350e2b)。

## [1.0.1] — Git 記録日: 2026-01-25

- パッケージ専用の版指定を共通の Version 指定へ変更し、アセンブリとパッケージのバージョン管理を整合。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/b8430cbcc6b50f89ccac25d91167819b27c05c22) / [変更差分](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/compare/a215e0dfbe91c3dde003347386012c93af082c54...b8430cbcc6b50f89ccac25d91167819b27c05c22)。

## [1.0.0] — Git 記録日: 2026-01-25

- 1llum1n4t1s.AlphaFS フォークのパッケージ情報・参照先を設定。

出典: [版の記録](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS/commit/a215e0dfbe91c3dde003347386012c93af082c54)。

---

Version 2.3  (2018-XX-XX)
-------------

### New Features

- Issue #451: Add overloaded method for `File.GetSize` to retrieve the size of all streams.  
- Issue #454: Add method `Directory.GetSize` to return the size of all alternate data streams of the specified directory and its files.
- Issue #464: Add overloaded methods for `Directory.Copy/Move` supporting `DirectoryEnumerationFilters`.
- Issue #465: Add overloaded methods for `File.Copy/Move` supporting retry.
- Issue #466: Add properties `ErrorRetry` and `ErrorRetryTimeout` to `DirectoryEnumerationFilters` class.
- Issue #467: Add property `CopyOptions.CopyTimestamp`.

### Improvements

- Issue #426: Correct casing of enum `STREAM_ATTRIBUTES`
- Issue #459: Modify method `Directory.CreateDirectoryCore` to return `null` as well as `DirectoryInfo` instance.
- Issue #461: Remove text `:$DATA` from `AlternateDataStream.FullPath` property.
- Issue #462: Add `IEquatable<T>` to applicable info classes.
- Issue #463: Add `[Serializable]` attribute to applicable info classes.
- Issue #470: Change AlphaFS implementations of method `DirectoryInfo.Create` to return `DirectoryInfo` instance instead of void.
- Issue #471: Add overloaded method `Directory.CountFileSystemObjects` supporting `DirectoryEnumerationFilters`.
- Issue #472: Add method `DirectoryInfo.ExistsJunction`.
- Issue #473: Change AlphaFS implementations of method `File.Copy` to return `CopyMoveResult` instance instead of `void`.
- Issue #475: Prevent `File.SetFsoDateTimeCore` from throwing `UnauthorizedAccessException`.
- Issue #477: Move method `Path.GetMappedConnectionName` to `Host` class.
- Issue #478: Move method `Path.GetMappedUncName` to `Host` class.
- Added missing overloaded methods regarding timestamps and symbolic links.
- Marked `Directory.Copy`/`DirectoryInfo.Copy` methods using parameters `overwrite` and `preserveDates` as obsolete. Use one of the `CopyOptions`.
- Fixed a `Directory.Move` unit test creating endless drive mappings on error.
- Issue #504: Move to Cake build system.
- Issue #502: Upgrade to MSTest v2.
- Issue #501: Documentation generated using DocFX.
- Issue #482: AlphaFS .NET Core compatibility (netstandard 2.0 support added)

### Breaking Changes

- Issue #426: Correct casing of enum `STREAM_ATTRIBUTES`
- Issue #461: Remove text `:$DATA` from `AlternateDataStream.FullPath` property.
- Issue #477: Move method `Path.GetMappedConnectionName` to `Host` class.
- Issue #478: Move method `Path.GetMappedUncName` to `Host` class.
- Issue #500: Drop support for .NET framework versions prior to .NET 4.5

Version 2.2.6  (2018-08-18)
-------------

### Bugs Fixed

- Issue #488: `Path.CheckInvalidPathChars` in `Path.Helpers.cs` should be case insensitive  (Thx GuyTe)
- Issue #489: `File.Copy` fails with `DirectoryNotFoundException` on long path  (Thx CyberSinh)

### Improvements

- Issue #487: Ensure replace is done case-insensitive  (Thx Genbox)


Version 2.2.5  (2018-07-27)
-------------

### Bugs Fixed

- Issue #479: `File.Move` on a file opened with `FileShare.Delete` succeeds but throws `IOException`.  (Thx oguimbal)
- Issue #480: `Directory.Delete(, true, true)` occasionally throws `DirectoryNotEmptyException`.


Version 2.2.4  (2018-07-12)
-------------

### Bugs Fixed

- Issue #468: Method `Directory.CopyTimestamps` should throw `DirectoryNotFoundException`.
- Issue #469: Method `Directory.GetFileIdInfo` should throw `DirectoryNotFoundException`.
- Issue #474: Method `Directory.EnumerateAlternateDataStreams` should throw `DirectoryNotFoundException`.
- Issue #476: Method `Directory.GetChangeTime` should throw `DirectoryNotFoundException`.


Version 2.2.3  (2018-06-14)
-------------

### Bugs Fixed

- Issue #456: Methods `Directory/File.Copy` throw `DeviceNotReadyException` when using `GLOBALROOT` source path.  (Thx VladimirK69)
- Issue #457: `FileInfo.Exists` is `true` when instance is created from a directory path.
- Issue #458: `Directory.Copy` sometimes does not create the file's parent folder, throwing `System.IO.DirectoryNotFoundException`.

### Improvements

- Added unit tests for GLOBALROOT source path so that it can never happen again!


Version 2.2.2  (2018-06-05)
-------------

### Bugs Fixed

- Issue #434: `Directory.Move` operation worked in v2.0.1, but now fails in v.2.2.1  (Thx warrenlbrown)
- Issue #436: `Directory.GetFiles()` with relative path  (Thx stellarbear)
- Issue #437: Fixed `PathTooLongException` for boundary case of directory name length in `Path.NormalizePath`  (Thx okrushelnitsky)
- Issue #441: `InvalidOperationException` on `Directory.EnumerateFileSystemEntries()` (Thx CyberSinh)
- Issue #444: Exception when moving or renaming a folder after updating from 2.1.3 to 2.2.1  (Thx mlaily)

### Improvements

- Issue #149: Split unit-tests.
- Fixed all Sandcastle Help File Builder warnings.
- Because of bug fixes, the correct source-/destination folder is now shown in exceptions thrown from Directory/File.Copy/Move methods, instead of always showing the source folder path.
- Improved some unit tests that would fail when a removable drive is already removed but there's still a cached reference.


Version 2.2.1  (2018-04-05)
-------------

### Bugs Fixed

- Issue #433: Directory.ExistsDriveOrFolderOrFile fails on global root path, so a simple file copy from a shadow copy fails with "device GLOBALROOT not ready" (Thx CyberSinh)


Version 2.2  (2018-03-25)
-----------

### Bugs Fixed

- Issue #268: There are multiple warnings when building the documentation.
- Issue #286: Property `FileSystemEntryInfo.AlternateFileName` is always an empty string.
- Issue #292: `CopyOptions.CopySymbolicLink` not working in 2.1.2  (Thx v2kiran)
- Issue #325: `DeleteEmptySubdirectories` (with `recursive=true`) throws `System.IO.DirectoryNotFoundException`  (Thx kryvoplias)
- Issue #328: Several instances of `ArgumentException.ParamName` not set/used correctly  (Thx elgonzo)
- Issue #330: Correct the parameter order for Privilege class constructors using the `ArgumentNullException`.
- Issue #339: `Directory/File.Encrypt/Decrypt` should restore read-only attribute.
- Issue #340: `DirectoryReadOnlyException` inherits from `System.IO.IOException`, wrong?
- Issue #344: `Directory.Copy` throws `UnauthorizedAccessException` "The target file is a directory, not a file", while it is a file.
- Issue #349: `File.GetFileSystemEntryInfoCore` should throw `Directory/FileNotFoundException`, depending on `isFolder` argument.
- Issue #369: `Directory.EnumerateFileSystemEntryInfos` does not return subdirectories with spaces as name.  (Thx Lupinho)
- Issue #371: Fix `.gitignore` to accommodate new directory structure in AlphaFS.UnitTest project.  (Thx damiarnold)
- Issue #372: `SetFsoDateTimeCore` should always use `BackupSemantics`.  (Thx damiarnold)
- Issue #374: Initializing `DriveInfo` instance with only a letter throws `System.ArgumentException`.
- Issue #375: What is the expected result of `Path.LocalToUnc()` ?  (Thx damiarnold)
- Issue #376: `Path.LocalToUnc(path, true)` does not return UNC path in long path form.  (Thx damiarnold) 
- Issue #379: `Path.LocalToUnc()` isn't handling trailing directory separators for mapped drives.  (Thx damiarnold)
- Issue #381: Change property `ByHandleFileInfo.VolumeSerialNumber` from `int` to `long`.
- Issue #386: `Network.Host.EnumerateDrives()` returns invalid data.
- Issue #400: `Directory.CopyDeleteCore` creates destination folder when source folder does not exist.
- Issue #412: Method `Volume.GetVolumeDeviceName` returns wrong result.
- Issue #417: Using a file opened in append mode will fail after a gc occurs  (Thx elgonzo)
- Issue #417: `File.OpenRead` method overloads do not use `FileShare.Read`  (Thx elgonzo)
- Issue #427: `System.IO.FileSystemInfo.Refresh()` is public; but AlphaFS `FileSystemInfo.Refresh()` is protected  (Thx elgonzo)

### New Features

- Issue #212: Provide a way to retrieve errors when you choose to `ContinueOnException`
- Issue #314: Added Feature: `Directory.GetFileSystemEntryInfo`  (Thx besoft)
- Issue #322: Search files/folders using multiple filters. (Thx besoft)
- Issue #336: Implement methods for `Directory` class: `CreateJunction`, `DeleteJunction` and `ExistsJunction`.
- Issue #338: Add convenience method `Directory.IsEmpty`
- Issue #342: Add instance method: `FileInfo.IsLocked()`
- Issue #343: Add method `File.GetProcessForFileLock`
- Issue #347: Implement method `Directory.CreateSymbolicLink`
- Issue #348: Implement method `Directory.GetLinkTargetInfo`
- Issue #351: Enable copying of Directory symbolic links.
- Issue #353: Modify method `Directory.GetFileSystemEntryInfo` to return `FileSystemEntryInfo` structure for directories supporting also root directories, e.g., `C:\`  (Thx besoft)
- Issue #354: Add methods `File.GetFileId` and `Directory.GetFileId` to return a unique file identifier.  (Thx besoft)
- Issue #370: Relative path from a full path  (Thx QbProg)
- Issue #373: Improve method `Directory.CreateDirectory` to allow creating a folder consisting only of spaces.
- Issue #414: Add additional `Network.Host` methods.
- Issue #415: Added `ProcessContext` static class to determine the context of the current process.
- Issue #422: Add `Copy-FileWithProgress.ps1` demonstrating file copy with progress report.
- Issue #423: Add `Copy-DirectoryWithProgress.ps1` demonstrating directory copy with progress report.

### Improvements

- Issue #273: Improve methods `Directory/File.CopyMoveCore`: Make code recursive-aware, skip additional path checks and validations.
- Issue #274: Improve methods `Directory/File.CopyMoveCore`: Improve detection of same volume.
- Issue #275: Improve methods `Directory/File.CopyMoveCore`: Eliminate recursion.
- Issue #277: `Directory.DeleteDirectoryCore()`: Eliminate recursion.
- Issue #278: `Directory.DeleteEmptySubdirectoriesCore()`: Eliminate recursion.
- Issue #303: `Path.Constants.cs`: Don't use `CurrentCulture`  (Thx HugoRoss)
- Issue #306: Include `ShareInfoLevel.Info502` and set as a fallback in `GetShareInfoCore()`  (Thx damiarnold)
- Issue #326: Add parameter `bool preserveDates` and created overloaded `Directory.Copy` methods to support this.
- Issue #331: Rename method `File/Directory.TransferTimestamps` to `CopyTimestamps`.
- Issue #335: Add overloaded methods to `File/Directory.TransferTimestamps` to apply to ReparsePoint.
- Issue #341: Improve usage of method `NativeError.ThrowException` and `Marshal.GetLastWin32Error`
- Issue #345: `AlreadyExistsException` should only throw message from 1 error.
- Issue #350: Add overloaded methods `Directory.GetFileSystemEntryInfo`
- Issue #352: Ignore `NonInterpretedPathPrefix` in methods: `Path.GetFullPathCore` and `Path.GetLongPathCore`  (Thx besoft)
- Issue #355: Methods throwing an `IOException` due to error code 17 (`ERROR_NOT_SAME_DEVICE`) now throw a specific exception (`NotSameDeviceException`)
- Issue #357: Added new Windows 10 property: `DirectAccess` (Win32 `FILE_DAX_VOLUME`) to `VolumeInfo` class.
- Issue #360: Add overloaded `Directory.EnumerateDirectories` methods that support `DirectoryEnumerationFilters`.  (Thx SignFinder)
- Issue #364: Avoid unnecessary allocations from Random construction in FileSystemInfo  (Thx danmosemsft)
- Issue #377: Rename enum member: `OperatingSystem.EnumOsName.WindowsServer` to: `OperatingSystem.EnumOsName.WindowsServer2016`
- Issue #378: `DiskSpaceInfo` should respect `CultureInfo.CurrentCulture` for number formatting.
- Issue #385: Correct applied fileSystemRights operator in method `File.Create()`.
- Issue #387: Replace `handle.IsInvalid` checks with a call to method `NativeMethods.IsValidHandle()`.
- Issue #388: Change method `Network.Host.EnumerateDrives()` return type from `string` to `DriveInfo`.
- Issue #394: Improve percentage output of properties `DiskSpaceInfo`- `AvailableFreeSpacePercent` and `UsedSpacePercent`
- Issue #401: CreateDirectory `ERROR_ACCESS_DENIED` reports parent folder.
- Issue #402: Remove long path prefix from `NativeError.ThrowException` messages with paths.
- Issue #408: Add `FileSystemEntryInfo.Extension` string property.
- Issue #416: Rename property `DeviceInfo.Class` to `DeviceInfo.DeviceClass`

### Breaking Changes

- Issue #331: Rename method `File/Directory.TransferTimestamps` to `CopyTimestamps`. Currently non-breaking, the old methods are still there.
- Issue #340: `DirectoryReadOnlyException` inherits from `System.IO.IOException`, wrong?
- Issue #350: Add overloaded methods `Directory.GetFileSystemEntryInfo`. Current code to retrieve a directory using `File.GetFileSystemEntryInfo` will now fail. Use `Directory.GetFileSystemEntryInfo` instead.
- Issue #377: Rename enum member: `OperatingSystem.EnumOsName.WindowsServer` to: `OperatingSystem.EnumOsName.WindowsServer2016`
- Issue #381: Change property `ByHandleFileInfo.VolumeSerialNumber` from `int` to `long`.
- Issue #388: Change method `Network.Host.EnumerateDrives()` return type from `string` to `DriveInfo`.
- Issue #391: Mark AlphaFS enumerating methods that use both `searchPattern` and `DirectoryEnumerationFilters` as obsolete.
- Issue #416: Rename property `DeviceInfo.Class` to `DeviceInfo.DeviceClass`


Version 2.1.3 (2017-06-05)
-------------

### Bugs Fixed

- Issue #288: `Directory.Exists` on root drive problem has come back with recent updates  (Thx warrenlbrown)
- Issue #289: `Alphaleonis.Win32.Network.Host.GetShareInfo` doesn't work since 2.1.0  (Thx Schoolmonkey/damiarnold)
- Issue #296: Folder rename (casing) throws IOException with HResult `ERROR_SAME_DRIVE`  (Thx doormalena)
- Issue #297: Incorrect domain returned from `Host.EnumerateDomainDfsRoot` when specifying domain  (Thx damiarnold)
- Issue #299: `FileInfo.MoveTo` and `DirectoryInfo.MoveTo` throw `ArgumentNullException` on empty destination path  (Thx doormalena)
- Issue #312: `Volume.EnumerateVolumes` skips first volume  (Thx springy76)
- Issue #313: `GetHostShareFromPath()` fails with spaces in share name  (Thx damiarnold)
- Issue #320: Minor changes in comments in `Win32Errors.cs` to eliminate compiler warnings.  (Thx besoft)
- Issue #321: `DirectoryInfo.CopyToMoveToCore()` calls `Path.GetExtendedLengthPathCore()` without `Transaction` parameter.


Version 2.1.2 (2016-10-30)
-------------

### Bugs Fixed

- Issue #270: Method `File.GetFileSystemEntryInfoCore` uses wildcard ? (questionmark) instead of * (asterisk)
- Issue #276: `Directory.DeleteDirectory()`: Method can get stuck in infinite loop.
- Issue #279: The unit tests for CRC32/64 are failing.


Version 2.1  (2016-09-29)
-----------

### New Features/Enhancements

- Issue #3: Added methods for backing up and restoring encrypted files:
	* `File.ImportEncryptedFileRaw`
	* `File.ExportEncryptedFileRaw`
	* `Directory.ImportEncryptedDirectoryRaw`
	* `Directory.ExportEncryptedDirectoryRaw`
- Issue #2  : Unit tests for methods: `File.OpenRead()`, `File.OpenText()` and `File.Replace()` are missing.
- Issue #101: The release now also contains a build targetting .NET 4.5.2.
- Issue #109: Add overloaded methods for `Host.EnumerateShares()`.
- Issue #112: Add `CreationTimeUtc`, `LastAccessTimeUtc` and `LastWriteTimeUtc` to "Info" classes.
- Issue #119: Fix `Path.IsLocalPath()` issues.
- Issue #125: AlphaFS is now CLSCompliant.
- Issue #127: Modify method `Volume.QueryDosDevice()` so that is doesn't rely on `Path.IsLocalPath()` anymore.
- Issue #130: Modify method `Path.LocalToUnc()` so that is doesn't rely on `Path.IsLocalPath()` anymore.
- Issue #131: Modify method `Path.GetPathRoot()` to handle UNC paths in long path format.
- Issue #132: Modify method `VolumeInfo()` constructor to better handle input paths.
- Issue #133: Add missing unit test `Host.GetHostShareFromPath()`.
- Issue #134: Improved upon `FindFileSystemEntryInfo.FindFirstFile()` when throwing `Directory-/FileNotFoundException()`.
- Issue #138: Modify `GetShareInfo()` to use `SafeGlobalMemoryBufferHandle` instead of `IntPtr`.
- Issue #139: Modify `GetDfsInfoInternal()` to use `SafeGlobalMemoryBufferHandle` instead of `IntPtr`.
- Issue #141: Remove obsolete Resources (resx) string messages.
- Issue #142: Move literal strings to Resources (resx).
- Issue #144: Add `DirectoryInfo.EnumerateXxx()` methods with support for `DirectoryEnumerationOptions` enum.
- Issue #151: Add `Directory.EnumerateXxx()` methods with support for `DirectoryEnumerationOptions`- and `PathFormat` enum.
- Issue #154: Modify private method `FindFileSystemEntryInfo.FindFirstFile()` to report the full path on Exception. 
- Issue #146: Add method `DirectoryInfo.EnumerateAlternateDataStreams()`.
- Issue #147: Add overloaded methods to set Reparse Point Timestamp.  (Thx rstarkov)
- Issue #150: Enhancement: `File.IsLocked()`
- Issue #158: Add SuppressUnmanagedCodeSecurity attribute to [DllImport] tag.
- Issue #184: `File.CreateSymbolicLink()` should throw `PlatformNotSupportedException()` if OS < Vista. 
- Issue #186: Replace WIN32 API `NativeMethods.GetVersionEx()` with `NativeMethods.RtlGetVersion()`.
- Issue #188: Make `ShareInfo` class property setters private: `ShareType`, `ResourceType`.
- Issue #189: Improve method `Utils.UnitSizeToText()`.
- Issue #190: Add overloaded methods for `File/Directory.Get/SetAccessControl()` that accept `SafeFileHandle`.
- Issue #191: Make class `BackupFileStream` sealed.
- Issue #192: Add `null`-checks to `SafeHandle.IsInvalid` usage.
- Issue #193: Use unicode version of WIN32 API `OpenEncryptedFileRaw()`.
- Issue #194: Add bitshift for Marshal.GetHRForException(ex) usage. 
- Issue #195: Add useful FileAttributes as properties to `FileSystemEntryInfo` class.
- Issue #199: Change `FindFileSystemEntryInfo.FindFirstFile()` to show actual path instead of inputpath on access error.
- Issue #214: Howto `Get-Filehash`.
- Issue #235: Implement unicode versions of methods: CM_Connect_Machine and CM_Get_Device_ID_Ex.
- Issue #239: Enable long path support for `File.CreateSymbolicLink()` source parameter.
- Issue #240: Add `KeepDotOrSpace` to `GetFullPathOptions` enum.
- Issue #241: Add method `Path.GetFullPath()` overload that supports `GetFullPathOptions` enum.
- Issue #245: Implement CRC-32/64 (Thanks to Damien Guard for implementing his code).
- Issue #247: Add method `FileInfo.GetHash()`.
- Issue #251: Implement unicode versions of `Directory.GetCurrentDirectory()` and `Directory.SetCurrentDirectory()`.
- Issue #266: Add PowerShell script: `Enumerate-FileSystemEntryInfos.ps1`

### Bugs Fixed

- Issue #50 : `Path.GetLongPath()` does not prefix on "C:\", should it?
- Issue #60 : Remove all use of "Problematic" methods such as `DangerousAddRef` and `DangerousGetHandle()`.
- Issue #160: `File.CreateSymbolicLink()` creates shortcut with no target.  (Thx martin-john-green)
- Issue #162: `File.AppendAllLines()` concatenates content into one line.  (Thx pavelhritonenko)
- Issue #166: `File.Exists` & `Directory.Exists` fail when path has leading space(s).
- Issue #168: Error on `File.Open()` with access-mode Append?
- Issue #169: `DirectoryInfo .ToString()` returns path with `\\UNC` prefix.
- Issue #176: At `DirectoryInfo.GetFileSystemInfos()`, Long path prefix of GLOBALROOT path is missing.  (Thx diontools)
- Issue #179: `Path.GetFileName()` with an empty string throws an exception.  (Thx brutaldev)
- Issue #180: Network connects methods hangs in Windows service when credentials fail.  (Thx brutaldev)
- Issue #181: `File.OpenWrite()` should create file if it doesn't exist.  (Thomas Levesque)
- Issue #183: Add `SafeFileHandle` null check for BackupFileStream.Dispose.  (Thx diontools)
- Issue #185: Correct pinvoke signatures of `CreateSymbolicLink()` and `CreateSymbolicLinkTransacted()` functions.
- Issue #196: Replace usage of `ExtendedFileAttributes.None` with `ExtendedFileAttributes.Normal`.
- Issue #197: Fix: Prevent normalization of GlobalRootPrefix paths.
- Issue #198: `Path.GetRegularPathCore()` should not normalize `\\?\Volume` prefix.
- Issue #201: Some exceptions contain an incorrect `HRESULT` (Thomas Levesque)
- Issue #203: `Directory.GetDirectories()` and `Directory.GetFiles()` return absolute paths when given relative argument.
- Issue #204: Giving empty string to `Directory.GetFileName()` and related methods throws exception.
- Issue #206: `File.GetLastWriteTime()` throws exception for non-existing path.
- Issue #217: `File.Replace()` raises an exception.
- Issue #218: `Volume.GetVolumeInfo()` fails for global root paths.
- Issue #219: Mismatching Implementation to `System.IO.Path.GetDirectoryName()`.
- Issue #226: `DirectoryInfo` using searchoption.
- Issue #232: Enable null for destinationBackupFileName for `File.Replace()` and `FileInfo.Replace()`.
- Issue #234: `Path.CheckInvalidPathChars` breaks `IsPathRooted` for whitespace strings.
- Issue #242: `File.Open(file, System.IO.FileMode.Append)` does not append.
- Issue #244: `File.Copy(src, dst, true)` does not respect `FILE_ATTRIBUTE_READONLY`.
- Issue #246: Using `Directory.EnumerateFileSystemEntries()` recursively with a relative path may fail.
- Issue #248: `Directory.Move()` throws `FileNotFoundException` instead of `DirectoryNotFoundException` when source folder doesn't exist.
- Issue #249: Change `File.GetHashCore()` `.ToString("X2")` to `.ToString("X2", CultureInfo.InvariantCulture)`.
- Issue #252: Correct `FileSystemEntryInfos.FullPath` property when input path is a dot (current directory).
- Issue #253: Apply `Dispose()` to method `File.GetHashCore()`.
- Issue #254: Change `File.GetHashCore()` output from `.ToLowerInvariant()` to `.ToUpperInvariant()`.
- Issue #255: Creating Folder with Empty name. (ardestan)
- Issue #256: `Directory.Move()` not working over volumes with `MoveOptions.CopyAllowed`.  (Thx frontier777)
- Issue #263: `Directory.GetDirectories()` Method `(String, String, SearchOption)` with pattern "* "  (Thx ardestan)

### Breaking Changes

- Issue #113: Change names of time related properties on `FileSystemEntryInfo` to conform with `FileInfo/DirectoryInfo`.
- Issue #126: Suffix the name of all methods working with TxF with "Transacted".
- Issue #128: Remove `Path.IsLocalPath()` in favour of `Path.IsUncPath()`.
- Issue #140: Replace internal `DFS_INFO_4` structure with `DFS_INFO_9`.
- Issue #184: `File.CreateSymbolicLink()` should throw `PlatformNotSupportedException()` if OS < Vista. 
- Issue #250: Change `FileSystemEntryInfo.ToString()` to show full path instead of `ReparsePointTag`.


Version 2.0.1  (2015-02-07)
-------------

### Bugs Fixed

- Issue #104: `VolumeInfo.Refresh()` fails with `System.IO.IOException`: (234)".
- Issue #108: `Volume.Refresh()` should throw `DeviceNotReadyException`.
- Issue #110: `Directory.GetDirectoryRoot()` should throw `System.ArgumentException`.
- Issue #117: Fix `Path.GetFullPath()` issues. 
- Issue #123: When `Directory.Encrypt/Decrypt()` is non-recursive, only process the folder.
- Issue #124: Unit tests for long/short path are failing.


Version 2.0  (2015-01-16)
-----------
* New: The public key of AlphaFS.dll has changed, delay-signing is no longer used.
* New: Unit Tests, also act as code samples.
* New: Numerous bugfixes, optimizations and (AlphaFS) overloaded methods implementations.
* New: Complete implementation of .NET 4.5 File(Info) and Directory(Info) classes.
* New: Complete implementation of .NET 4.5 DriveInfo() class and with UNC support.
* New: Complete implementation of .NET 4.5 Path() class.
* New: Implemented Unicode aka "Long Path" handling for all Win32 API functions that support it.
* New: Added support for NuGet.
* New: Added support for building against .NET 4.0, 4.5, and 4.5.1 in addition to 3.5.
* New: Supports networking by enumerating hosts and shares (SMB/DFS) and connect/disconnect to/from network resources (AlphaFS.Network.Host() class).
* New: Supports working with NTFS ADS (Alternate Data Streams) on files and folders (AlphaFS.Filesystem.AlternateDataStreamInfo() class).
* New: Supports enumerating connected PnP devices (AlphaFS.Filesystem.Device() / AlphaFS.Filesystem.DeviceInfo() classes).
* New: Supports extracting icons from files and folders (AlphaFS.Filesystem.Shell32Info() class).
* New: Supports PathFormat parameter for numerous methods to control path normalization. This speeds up things internally (less string processing and GetFullPath() calls) and also enables working with files and folders with a trailing dot or space:
	* `RelativePath` (slow): (default) Path will be checked and resolved to an absolute path. Unicode prefix is applied.
	* `FullPath`  (fast): Path is an absolute path. Unicode prefix is applied.
	* `LongFullPath`  (fastest): Path is already an absolute path with Unicode prefix. Use as is.
* Mod: Enabled KernelTransaction parameter for all Win32 API functions that support it.
* Mod: Added public read only properties to class FileSystemInfo(). Available for: DirectoryInfo() / FileInfo():
	* EntryInfo	 : Instance of the FileSystemEntryInfo() class.
	* Transaction  : Represents the KernelTransaction that was passed to the constructor.	
* Mod: Added more entries to enum ReparsePointTag.
* Mod: Removed method Directory.CountFiles() and added method Directory.CountFileSystemObjects().	
* Mod: Removed method Directory.GetFullFileSystemEntries() and added method Directory.EnumerateFileSystemEntryInfos().
	* Note: This new method currently does not support DirectoryEnumerationExceptionHandler, this will probably be added in a future release.
* Mod: Renamed method Directory.GetFileIdBothDirectoryInfo() to Directory.EnumerateFileIdBothDirectoryInfo().
* Mod: Method Directory.CreateDirectory() signature change: Using template directory. Ability for NTFS-compressed folders added.
* Mod: Method Directory.GetProperties() signature change.
* Mod: Renamed method File.GetFileInformationByHandle() to File.GetFileInfoByHandle().
* Mod: Removed overloaded method File.GetFileInformationByHandle(FileStream).h
* Mod: Removed overloaded AlphaFS methods File.Move() using MoveFileOptions and CopyProgressRoutine, and added method File.Move().
* Mod: Renamed method Volume.GetDeviceForVolumeName() to Volume.GetVolumeDeviceName().
* Mod: Renamed method Volume.GetDisplayNameForVolume() to Volume.GetVolumeDisplayName().
* Mod: Renamed method Volume.GetVolumeInformation() to Volume.GetVolumeInfo().
* Mod: Renamed method Volume.GetVolumeMountPoints() to Volume.EnumerateVolumeMountPoints().
* Mod: Renamed method Volume.GetVolumePathNamesForVolume() to Volume.EnumerateVolumePathNames().
* Mod: Renamed method Volume.GetVolumes() to Volume.EnumerateVolumes().
* Mod: Method Volume.DefineDosDevice() signature change.
* Mod: Method Volume.QueryDosDevice() signature change.
* Mod: Method Volume.QueryAllDosDevices() signature change.
* Mod: Removed method Volume.GetLogicalDrives() in favor of method Directory.GetLogicalDrives().
* Mod: Class VolumeInfo() constructor signature change.
* Mod: Class VolumeInfo() properties updated/changed.
* Mod: Added method Volume.Refresh().
* Mod: Changed struct DiskSpaceInfo() to class.
* Mod: Class DiskSpaceInfo() constructor signature change.
* Mod: Class DiskSpaceInfo() properties added.
* Mod: Added method DiskSpaceInfo.Refresh().
* Mod: Refactored Path() class.
* Mod: Improved upon the correct (.NET) exceptions thrown. Added AlphaFS specific: DirectoryReadOnlyException and FileReadOnlyException.
* Removed classes PathInfoXxx().
* Removed method Path.IsValidPath(), was part of PathInfo() class.
* Removed IllegalPathException.
* Removed enum DriveType in favor of System.IO.DriveType enum.
* Removed enum FileAccess in favor of System.IO.FileAccess enum.
* Removed enum FileAttributes in favor of System.IO.FileAttributes enum.
* Removed enum FileMode in favor of System.IO.FileMode enum.
* Removed enum FileOptions in favor of System.IO.FileOptions enum.
* Removed enum FileShare in favor of System.IO.FileShare enum.
* Removed enum FileSystemRights in favor of System.Security.AccessControl.FileSystemRights enum.
* Removed enum FileType, obsolete.
* Removed enum EnumerationExceptionDecision, obsolete.
* Removed enum IoControlCode.cs, obsolete.
* Renamed enum CopyProgressResult to CopyMoveProgressResult.
* Renamed enum MoveFileOptions to MoveOptions.
* Renamed class DeviceIo to Device.
* Renamed delegate CopyProgressResult to CopyMoveProgressResult.


Version 1.5  (2014-05-20)
-----------
   * New: Various file system objects enumeration methods in Directory class.
   * Numerous bugfixes and optimizations
   * New: more unit tests
   * New: VS 2010 help file format, aka Help Viewer 1, dumped MS HELP 2 format


Version 1.0
-----------
  * New: Directory.GetFileIdBothDirectoryInfo, which provides access to the GetFileInformationByHandleEx Win32 API 
         function with the FileInformationClass set to FileIdBothDirectoryInfo.
  * New: Directory.CountFiles
  * Mod: Additional overloads for File.Open method.
  * Mod: FileAttributes.Invalid flag removed.
  * New: Directory.GetProperties method for retrieving aggregated information about files in a directory.
  * New: File.GetFileInformationByHandle added providing information about file index and link count.
  * New: KernelTransaction can now be created from a System.Transaction to participate in the ambient transaction
  * New: File.GetHardlinks providing an enumeration about all hardlinks pointing to the same file.
  * Mod: Many improvements and bug-fixes to Path/PathInfo path-parsing.
  * Mod: More functions for manipulating timestamps on files and directories.
  * Mod: Directory.GetFullFileSystemEntries added to provide more convenient usage of the FileSystemEnumerator.
  * Mod: ... and many more minor changes and fixes.

Version 0.7 alpha
-----------------
  * New: DirectoryInfo and FileInfo classes added
  * New: PathInfo.GetLongFullPath() and Path.GetLongFullPath() methods added
  * Mod: Path and PathInfo got many bugfixes, and some new functionality was added.
  * Mod: AlphaFS now targets the .NET Framework 2.0 instead of 3.5 previously.
  * Mod: KernelTransaction can now be created from, and participate in a System.Transactions.Transaction.
  * New: BackupFileStream added, in support of the BackupWrite(), BackupRead() and BackupSeek() functions from the Win32 API.
  * Mod: Inheritance structure for several classes was modified, mainly to add MarshalByRefObject to the relevant classes.
  * Mod: FileSystemEntryInfo was changed to a reference type (class) instead of the previous value type (struct).
  * Mod: PathInfo now accepts more types of internal paths, such as \\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy5\ etc.
  * ... and many minor changes and fixes, not mentioned here.


Version 0.3.1
-------------
  * New: Added support for hardlinks and symbolic links in File.
  * New: Added Directory.EnableEncryption() and Directory.DisableEncryption()
  * New: Added File.GetCompressedSize()
  * Mod: Applied CLSCompliant(false) to the assembly
  * Mod: Improved error reporting, and cleanup of internal class NativeError.
  

Version 0.3.0
-------------
  * Initial release
