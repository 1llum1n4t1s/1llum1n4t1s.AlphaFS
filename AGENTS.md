# CLAUDE.md

This file provides guidance to Claude Code and other coding agents working in this repository.

## Project Overview

1llum1n4t1s.AlphaFS is a maintained fork of the archived [alphaleonis/AlphaFS](https://github.com/alphaleonis/AlphaFS) library. It is a .NET library providing extended Win32 file system functionality beyond `System.IO`, including extended-length paths (up to 32,000 chars), NTFS alternate data streams, junctions/hardlinks, transactional file operations (TxF), and SMB/DFS network access.

- **Target:** .NET 10.0-windows; the library is AnyCPU, the test host is x64, and the library is NativeAOT compatible (`IsAotCompatible=true`)
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

Tests use **MSTest 4.x** with Microsoft.Testing.Platform (MTP); the exact package versions are defined in `tests/AlphaFS.UnitTest/AlphaFS.UnitTest.csproj`. Many tests require elevated privileges or specific NTFS/network configurations, so some may skip or fail in non-privileged environments — treat those as environment-dependent rather than regressions.

Two environment variables change which tests run:

| Variable | Default | Effect |
|---|---|---|
| `ALPHAFS_SKIP_NETWORK_TESTS` | unset (network tests run) | Set to any value other than `0`/`false` to make the network half of each test `Assert.Inconclusive`. Used by CI, where no SMB admin share exists. The local half still runs first, so regressions are still caught. |
| `ALPHAFS_ENABLE_MACHINE_STATE_TESTS` | unset (test is skipped) | Opt-in for tests that permanently modify machine state. Currently gates `AlphaFS_DirectoryInfo_MoveTo_DelayUntilReboot_Local_Success`, which writes to `HKLM` (`PendingFileRenameOperations`). Leave unset unless you intend that change. |

Note that elevation is **not** an opt-out: tests guarded by `RequireElevation` skip only in *non*-elevated processes. GitHub runners are elevated, so those tests run in CI even when they are skipped on a local non-elevated shell.

## Architecture

The current component boundaries, operation flows, invariants, and design decisions are maintained in [DESIGN.md](DESIGN.md). Read it before changing public filesystem APIs, path normalization, native/COM resource lifetime, transactional operations, or release boundaries.

## Breaking Changes from Upstream

This fork modernized COM interop for NativeAOT. Key API changes:
1. `NetworkConnectionInfo.NetworkInfo` property → `GetNetworkInfo()` method
2. `NetworkInfo`, `NetworkConnectionInfo`, `Shell32Info` now implement `IDisposable` — use `using` statements (full list including `Shell32Info` is in [README.md §"Breaking Changes from Upstream"](README.md))

## CI/CD

NuGet publishing is triggered by pushes to `release/**` branches via `.github/workflows/publish.yml`. The workflow restores, builds, tests, packs, obtains a short-lived credential through NuGet.org Trusted Publishing, and pushes the package. Do not create or store a long-lived NuGet API key.
