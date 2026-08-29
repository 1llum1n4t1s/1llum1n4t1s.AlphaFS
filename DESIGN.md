# 1llum1n4t1s.AlphaFS Design

## Purpose and scope

1llum1n4t1s.AlphaFS is a Windows-only .NET library that exposes Win32 filesystem, volume, security, transaction, SMB, and DFS capabilities through APIs shaped like `System.IO`. It preserves AlphaFS features such as extended-length paths, alternate data streams, links, native backup streams, and TxF while maintaining the archived upstream code for .NET 10 and NativeAOT consumers.

The library targets `net10.0-windows`. Its boundary is the Windows API: platform-independent filesystem abstraction and non-Windows implementations are outside this repository.

## Components and responsibilities

| Component | Location | Responsibility and boundary |
| --- | --- | --- |
| Filesystem API | `src/AlphaFS/Filesystem/` | Public `File`, `Directory`, `Path`, `FileInfo`, and `DirectoryInfo` surfaces; path handling; enumeration; copy/move; links; streams; transactions; Shell integration. |
| Device and volume API | `src/AlphaFS/Device/` | Drives, volumes, disk-space data, devices, and volume mount operations. These files intentionally use the `Alphaleonis.Win32.Filesystem` namespace. |
| Network API | `src/AlphaFS/Network/` | `Host`, SMB/DFS queries, mapped connections, and Network List Manager COM projections. |
| Security API | `src/AlphaFS/Security/` | Token privileges, access-control support, security-descriptor interop, and CRC utilities. |
| Native boundary | `src/AlphaFS/**/Native Methods/` | Domain-specific P/Invoke and COM declarations. Higher-level APIs do not duplicate native declarations. |
| Lifetime and errors | `src/AlphaFS/Safe Handles/`, `src/AlphaFS/NativeError.cs` | Ownership of native handles and mapping Win32 status codes to managed exceptions. |
| Verification | `tests/AlphaFS.UnitTest/` | MSTest coverage for local, network, elevated, NTFS-specific, and machine-state-dependent behavior. |

The solution contains one packable library project and one x64 test project. The library itself remains AnyCPU so it can load in both x64 and Arm64 native .NET processes.

## API structure

- Static facades (`File`, `Directory`, and `Path`) mirror familiar `System.IO` entry points.
- Info objects (`FileInfo`, `DirectoryInfo`, volume and network information types) retain paths or native metadata for object-oriented use.
- Large public types are partial classes split by operation family. Shared core methods own normalization, native calls, error handling, and common option processing.
- Normal and transactional overloads converge on the same core path. Transactional calls carry a `KernelTransaction` and its safe handle to the corresponding TxF native API.

The API is System.IO-shaped rather than an exact clone. Intentional differences and fork-specific breaking changes are documented in [README.md](README.md#breaking-changes-from-upstream-alphaleonisalphafs).

## Operation flows

### Filesystem operation

1. A static facade or Info object receives a path and operation options.
2. Shared core code validates and normalizes the path, including UNC and `\\?\` extended-length forms. Recursive copy/move also rejects a destination at or below its source before traversal begins.
3. A domain-specific `NativeMethods` declaration invokes the Windows API; transactional overloads also pass the kernel transaction handle.
4. Tree copy/move preserves requested reparse points as links instead of traversing their targets. Emulated moves retain the source until the destination has completed successfully.
5. Safe-handle types own returned handles. Managed result types project native structures where an object result is required.
6. `NativeError` maps failed Win32 results to the library's managed exception contract.

### Network and Shell operation

1. `Host` or Shell-facing APIs call Win32 networking functions or COM wrappers.
2. Native or COM data is projected into `NetworkInfo`, `NetworkConnectionInfo`, `Shell32Info`, or related managed types. Security descriptors exposed by `DfsInfo` and `ShareInfo` are copied into owned memory before their NetAPI result buffers are released.
3. Types that retain COM references implement `IDisposable`; callers release them with `using`.

### Build and publication

1. `Directory.Build.props` is the single product-version source for both projects.
2. `global.json` selects Microsoft.Testing.Platform for `dotnet test`.
3. Pushes to `release/**` run restore, Release build, tests, and pack on Windows.
4. The workflow requires the release branch suffix to equal the package version before NuGet Trusted Publishing obtains a short-lived credential and uploads the package.

## Invariants

- Extended paths support up to `MaxPathUnicode` (32,700 characters); standard and extended path forms must not be mixed incorrectly.
- Path normalization deliberately remains stricter than `System.IO` for malformed or extended paths.
- Recursive copy/move never uses a destination at or below its source, and reparse points selected for link copying are not traversed.
- An emulated move deletes its source only after the destination is complete; replacement rollback must not leave the caller without either the original source or a valid destination.
- Native handles are represented by SafeHandle-derived owners. Native or COM ownership must not be replaced with untracked raw lifetime management.
- Managed network result objects never retain borrowed pointers into NetAPI enumeration buffers after those buffers are freed.
- `NetworkInfo`, `NetworkConnectionInfo`, and `Shell32Info` retain COM resources and therefore remain disposable.
- Transactional and non-transactional overloads share behavior except for the transaction boundary.
- Default `SetAccessControl` overloads write DACL changes; overloads with `AccessControlSections` are required for owner, group, or SACL writes. `BackupFileStream` keeps its all-sections backup behavior.
- `System.Security.Permissions` remains a direct dependency because `NativeError` exposes `System.Security.Policy.PolicyException` for `ERROR_BAD_RECOVERY_POLICY` through that package's type forwarding.
- The library project has no fixed `PlatformTarget`; the test host is x64 because several tests exercise Windows-native behavior under that architecture.
- Tests that need SMB shares, elevation, NTFS features, or permanent machine-state changes declare those environmental requirements. `ALPHAFS_SKIP_NETWORK_TESTS` skips only network halves, and `ALPHAFS_ENABLE_MACHINE_STATE_TESTS` is an explicit opt-in.
- Release publication accepts only `release/x.y.z` when `Directory.Build.props` contains the same `x.y.z`; NuGet credentials remain short-lived through Trusted Publishing.

## Adopted design decisions

| Decision | Reason | Trade-off |
| --- | --- | --- |
| Preserve System.IO-shaped facades | Existing AlphaFS callers can migrate with small namespace-level changes. | Newer `System.IO` members and return types cannot always be matched exactly. |
| Split large types into feature-based partial files | Keeps hundreds of filesystem operations navigable without changing their public type identity. | A complete behavior trace can cross several files and shared core methods. |
| Normalize paths before native calls | Centralizes long-path, UNC, and validation behavior across public overloads. | AlphaFS can reject malformed paths that current `System.IO` handles differently. |
| Use P/Invoke and SafeHandle ownership | Exposes Windows-only capabilities while making native cleanup deterministic and exception-safe. | The implementation is platform-specific and requires Windows-focused tests. |
| Stage destructive directory replacement and defer source deletion | Keeps emulated cross-volume moves recoverable when copying or replacement fails partway through. | Replacement needs temporary sibling paths and additional cleanup logic. |
| Make COM-backed projections disposable | NativeAOT-compatible wrappers can release retained COM references predictably. | This is a breaking lifetime contract compared with archived upstream APIs. |
| Keep the library AnyCPU and tests x64 | Consumers can load the library in x64 or Arm64 processes while native tests retain their established environment. | The local test host does not by itself validate every architecture. |
| Pin publishing Actions to commit SHAs | Protects the credential-bearing NuGet publication path from mutable tags. | Dependabot or manual maintenance must update both SHA and version comments. |
| Group dependency minor/patch updates but isolate majors | Reduces routine PR volume while keeping breaking changes independently reviewable. | Monthly grouped updates can delay individual patch adoption until the scheduled run. |

## Verification boundary

Required developer commands, environment-dependent test behavior, and packaging commands are maintained in [AGENTS.md](AGENTS.md). User-visible capabilities and compatibility notes are maintained in [README.md](README.md).
