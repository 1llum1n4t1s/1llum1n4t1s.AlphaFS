# 1llum1n4t1s.AlphaFS

> **これは [AlphaFS](https://github.com/alphaleonis/AlphaFS) のフォークです。**  
> 元リポジトリは 2024年12月にアーカイブされているため、本リポジトリで開発・保守を継続しています。  
> **フォーク元:** [alphaleonis/AlphaFS](https://github.com/alphaleonis/AlphaFS) → [1llum1n4t1s/1llum1n4t1s.AlphaFS](https://github.com/1llum1n4t1s/1llum1n4t1s.AlphaFS)

1llum1n4t1s.AlphaFS is a .NET library providing more complete Win32 file system functionality to the .NET platform than the standard `System.IO` classes.

## Introduction

The file system support in .NET is pretty good for most uses. However there are a few shortcomings, which this library tries to alleviate. The most notable deficiency of the standard .NET `System.IO` is the lack of support of advanced NTFS features, most notably extended length path support (eg. file/directory paths longer than 260 characters: `System.IO.PathTooLongException`).

### Feature Highlights

* Support for extended length paths (longer than 260 characters)
* Creating Junctions/Hardlinks
* Accessing hidden volumes
* Enumeration of volumes
* Transactional file operations
* Support for NTFS Alternate Data Streams (files/folders)
* Accessing network resources (SMB/DFS)
* Create and access folders/files that have leading/trailing space(s) in their name
* Folder/file enumerator supporting custom filtering and error reporting/recovery (access denied exceptions)
* ...and much more!

## What does AlphaFS provide?

AlphaFS provides a namespace (`Alphaleonis.Win32.Filesystem`) containing a number of classes. Most notable
are replications of the `System.IO.Path`, `System.IO.File`, `System.IO.FileInfo`, `System.IO.Directory` and `System.IO.DirectoryInfo`, all with support for the extended-length paths (up to 32.000 chars), full UNC support,
recursive file enumerations, native backups and manipulations with advanced flags and options.
They also contain extensions to these, and there are many more features for several functions.

When only  these `System.IO` classes are used, it is mostly a matter of replacing `using System.IO;`
with `using Alphaleonis.Win32.Filesystem;`.

### Where the drop-in replacement is not exact

`System.IO` has grown since AlphaFS was first written, so a few members now differ. All of the
differences below surface as compiler errors rather than silent misbehaviour, but they are worth
knowing before you swap the `using`:

| Member | `System.IO` | AlphaFS |
|---|---|---|
| `File.CreateSymbolicLink(path, target)` | returns `FileSystemInfo` | returns `void` |
| `Directory.CreateSymbolicLink(path, target)` | returns `FileSystemInfo` | returns `void` |
| `FileSystemInfo.LinkTarget` | property | not available — use `File.GetLinkTargetInfo(path)` |
| `File.ResolveLinkTarget(path, returnFinalTarget)` | returns `FileSystemInfo` | not available — use `File.GetLinkTargetInfo(path)` |

Note that the statement form (`File.CreateSymbolicLink(a, b);`) compiles and behaves identically on
both; only code that consumes the return value needs changing.

Another thing AlphaFS brings to the table is support for transactional NTFS (TxF). Almost every method in
these classes exist in two versions. One normal, and one that can work with transactions, more specifically the
kernel transaction manager. This means that file operations can be performed using the simple, lightweight KTM 
on NTFS file systems, through .NET, using the interface of the standard classes we are all used to.

AlphaFS also contains some NTFS security related functionality (in `Alphaleonis.Win32.Security`), providing 
the ability to enable token privileges for a user, which may be necessary for eg. changing ownership of a file.

The `Alphaleonis.Win32.Network` namespace together with the `Alphaleonis.Win32.Network.Host` class offers
network functionality to connect to SMB/DFS resources and easily access files and folders on the network,
all with extended-length paths support.

The library is Open Source, licensed under the MIT license.

## Breaking Changes from Upstream (alphaleonis/AlphaFS)

本フォークでは .NET 10 / NativeAOT 対応のために COM Interop 実装を刷新しました。
これに伴い、以下の公開 API に破壊的変更があります。

### 1. `NetworkConnectionInfo.NetworkInfo` プロパティ → `GetNetworkInfo()` メソッド

| Before (upstream) | After (this fork) |
|---|---|
| `networkConnection.NetworkInfo` | `networkConnection.GetNetworkInfo()` |

戻り値の `NetworkInfo` は `IDisposable` を実装しているため、`using` で囲んでください。

```csharp
// Before
var name = networkConnection.NetworkInfo.Name;

// After
using var networkInfo = networkConnection.GetNetworkInfo();
var name = networkInfo?.Name;
```

### 2. `NetworkInfo` が `IDisposable` を実装

内部で COM 参照（`NetworkWrapper`）を保持するようになったため、使用後は `Dispose()` の呼び出しが必要です。

### 3. `NetworkConnectionInfo` が `IDisposable` を実装

同様に内部で COM 参照（`NetworkConnectionWrapper`）を保持するため、`Dispose()` が必要です。
`Host.EnumerateNetworkConnections()` の列挙結果を使い終わったら Dispose してください。

### 4. `Shell32Info` が `IDisposable` を実装

内部で COM 参照（`QueryAssociationsWrapper`）を保持するようになったため、使用後は `Dispose()` の呼び出しが必要です。

```csharp
// Before
var info = new Shell32Info(path);

// After
using var info = new Shell32Info(path);
```

### 5. `System.IO` の現行挙動に合わせた変更

.NET Framework 時代の挙動を引き継いでいた箇所を、ドロップイン代替として現行の `System.IO` に合わせました。

| API | Before (this fork) | After |
|---|---|---|
| `Path.Combine("C:", "file")` | `C:file` (ドライブ相対) | `C:\file` |
| `DirectoryInfo.Parent` / `Root` / 列挙結果の `ToString()` | 名前のみ (`MyFolder`) | フルパス (`C:\dir\MyFolder`) |
| `FileInfo` / `DirectoryInfo` の `Create()` / `Delete()` | キャッシュを保持（`Refresh()` まで `Exists` が古い値） | 自動でキャッシュを無効化 |

`Path.GetFullPath` / `GetPathRoot` の拡張長パス (`\\?\`) 処理や、不正なパスに対する例外送出は
AlphaFS が意図的に `System.IO` より厳格なままです（本ライブラリの目的のため変更していません）。

### 6. `OperatingSystem.EnumOsName` に新しい値を追加

Windows 11 以降を `Later` として扱っていたのを改め、専用の値を割り当てました。
`Later` は「このライブラリが知らない将来の OS」だけを表します。既存の値は変更していません。

追加値: `WindowsServer2019` / `Windows11` / `WindowsServer2022` / `WindowsServer2025`
