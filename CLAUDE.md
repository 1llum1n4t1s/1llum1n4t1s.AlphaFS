# CLAUDE.md

This file provides guidance to Claude Code and other coding agents working in this repository.

## Project Overview

1llum1n4t1s.AlphaFS is a maintained fork of the archived [alphaleonis/AlphaFS](https://github.com/alphaleonis/AlphaFS) library. It is a .NET library providing extended Win32 file system functionality beyond `System.IO`, including extended-length paths (up to 32,000 chars), NTFS alternate data streams, junctions/hardlinks, transactional file operations (TxF), and SMB/DFS network access.

- **Target:** .NET 10.0-windows, x64, NativeAOT compatible (`IsAotCompatible=true`)
- **Language:** C# with `AllowUnsafeBlocks`
- **NuGet Package:** `1llum1n4t1s.AlphaFS`

## Build & Test Commands

```bash
# Build
rtk dotnet build AlphaFS.slnx

# Build Release
rtk dotnet build AlphaFS.slnx -c Release

# Run all tests
rtk dotnet test AlphaFS.slnx

# Run a specific test by name
rtk dotnet test tests/AlphaFS.UnitTest/AlphaFS.UnitTest.csproj --filter "FullyQualifiedName~TestMethodName"

# Pack NuGet package (outputs to artifacts/)
rtk dotnet pack src/AlphaFS/AlphaFS.csproj -c Release -o artifacts
```

Tests use **MSTest** (`MSTest.TestFramework` 4.3.2). Many tests require elevated privileges or specific NTFS/network configurations, so some may skip or fail in non-privileged environments — treat those as environment-dependent rather than regressions.

## Architecture

### Namespace → Directory Mapping

| Namespace | Source Directory | Purpose |
|---|---|---|
| `Alphaleonis.Win32.Filesystem` | `src/AlphaFS/Filesystem/` | File, Directory, Path, FileInfo, DirectoryInfo + extensions |
| `Alphaleonis.Win32.Network` | `src/AlphaFS/Network/` | Host class, SMB/DFS, network connections |
| `Alphaleonis.Win32.Security` | `src/AlphaFS/Security/` | Privilege elevation, CRC |
| `Alphaleonis.Win32` | `src/AlphaFS/Device/` | Volume, DriveInfo, DiskSpaceInfo, DeviceInfo |

### Key Design Patterns

**Partial classes split by feature** — Core classes like `Directory`, `File`, `DirectoryInfo`, `FileInfo` are split across many files by functionality (e.g., `Directory.Compress.cs`, `Directory.CopyMove.cs`, `Directory.CoreMethods.cs`). The `Directory Class/` and `File Class/` folders each contain 100+ files organized into subfolders by feature area.

**Static facades + Info classes** — Mirrors `System.IO` patterns: static `Directory`/`File`/`Path` classes alongside object-oriented `DirectoryInfo`/`FileInfo` wrappers. AlphaFS is designed as a drop-in replacement (`using Alphaleonis.Win32.Filesystem;` instead of `using System.IO;`).

**Dual normal/transactional APIs** — Most file operations have a transactional variant accepting a `KernelTransaction` parameter for TxF support.

**P/Invoke + SafeHandle wrappers** — All Win32 interop is in `Filesystem/Native Methods/NativeMethods.*.cs` (13 files split by domain). Safe handles in `Safe Handles/` (11 wrapper classes). Error mapping in `NativeError.cs`.

**COM Interop with IDisposable** — Shell32 and Network operations use COM wrappers. `Shell32Info`, `NetworkInfo`, and `NetworkConnectionInfo` all implement `IDisposable` to manage COM references (breaking change from upstream).

### Key Constants (NativeMethods.Constants.cs)

- `MaxPath = 260` (standard Windows limit)
- `MaxPathUnicode = 32700` (extended path support via `\\?\` prefix)
- `DefaultFileBufferSize = 65536` (file I/O buffer)
- `DefaultNativeQueryBufferSize = 4096` (native API scratch buffer)

## Breaking Changes from Upstream

This fork modernized COM interop for NativeAOT. Key API changes:
1. `NetworkConnectionInfo.NetworkInfo` property → `GetNetworkInfo()` method
2. `NetworkInfo`, `NetworkConnectionInfo`, `Shell32Info` now implement `IDisposable` — use `using` statements (full list including `Shell32Info` is in [README.md §"Breaking Changes from Upstream"](README.md))

## CI/CD

NuGet publishing is triggered by pushes to `release/**` branches via `.github/workflows/publish.yml`. The workflow: restore → build Release → pack → push to NuGet.
