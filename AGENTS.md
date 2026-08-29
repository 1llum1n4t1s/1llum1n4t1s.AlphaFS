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
dotnet build AlphaFS.slnx

# Build Release
dotnet build AlphaFS.slnx -c Release

# Run all tests
dotnet test AlphaFS.slnx

# Run a specific test by name
dotnet test tests/AlphaFS.UnitTest/AlphaFS.UnitTest.csproj --filter "FullyQualifiedName~TestMethodName"

# Pack NuGet package (outputs to artifacts/)
dotnet pack src/AlphaFS/AlphaFS.csproj -c Release -o artifacts
```

Tests use **MSTest** (`MSTest.TestFramework` 4.3.2). Many tests require elevated privileges or specific NTFS/network configurations, so some may skip or fail in non-privileged environments — treat those as environment-dependent rather than regressions.

Two environment variables change which tests run:

| Variable | Default | Effect |
|---|---|---|
| `ALPHAFS_SKIP_NETWORK_TESTS` | unset (network tests run) | Set to any value other than `0`/`false` to make the network half of each test `Assert.Inconclusive`. Used by CI, where no SMB admin share exists. The local half still runs first, so regressions are still caught. |
| `ALPHAFS_ENABLE_MACHINE_STATE_TESTS` | unset (test is skipped) | Opt-in for tests that permanently modify machine state. Currently gates `AlphaFS_DirectoryInfo_MoveTo_DelayUntilReboot_Local_Success`, which writes to `HKLM` (`PendingFileRenameOperations`). Leave unset unless you intend that change. |

Note that elevation is **not** an opt-out: tests guarded by `RequireElevation` skip only in *non*-elevated processes. GitHub runners are elevated, so those tests run in CI even when they are skipped on a local non-elevated shell.

## Architecture

### Namespace → Directory Mapping

| Namespace | Source Directory | Purpose |
|---|---|---|
| `Alphaleonis.Win32.Filesystem` | `src/AlphaFS/Filesystem/` | File, Directory, Path, FileInfo, DirectoryInfo + extensions |
| `Alphaleonis.Win32.Network` | `src/AlphaFS/Network/` | Host class, SMB/DFS, network connections |
| `Alphaleonis.Win32.Security` | `src/AlphaFS/Security/` | Privilege elevation, CRC |
| `Alphaleonis.Win32.Filesystem` | `src/AlphaFS/Device/` | Volume, DriveInfo, DiskSpaceInfo, DeviceInfo |
| `Alphaleonis.Win32` | `src/AlphaFS/` (repo root of the project) | `OperatingSystem`, `NativeError`, `Win32Errors`, `Resources`, and the `Safe Handles/` memory/token wrappers |

> Note: everything under `src/AlphaFS/Device/` declares `namespace Alphaleonis.Win32.Filesystem`, not `Alphaleonis.Win32`. Only the files listed in the last row live in the bare `Alphaleonis.Win32` namespace.

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

NuGet publishing is triggered by pushes to `release/**` branches via `.github/workflows/publish.yml`. The workflow restores, builds, tests, packs, obtains a short-lived credential through NuGet.org Trusted Publishing, and pushes the package. Do not create or store a long-lived NuGet API key.
