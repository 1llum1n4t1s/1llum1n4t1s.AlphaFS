# At-rest security review: path normalization and reparse safety

Review of 1llum1n4t1s.AlphaFS focusing on high-impact exploitable issues with complete attack chains (untrusted path input into a consuming app, real code path, impact beyond “the API does what it says”).

## Validated findings

### 1. Recursive copy/move “destination under source” guard is lexical-only and bypassable

- **Severity:** High
- **Location:** `src/AlphaFS/Filesystem/Directory Class/Directory Core Methods/Directory.CopyMoveCore.cs`
- **Evidence:**
  - Guard only when `isFolder && cma.IsCopy`: lines 60–63
  - Lexical prefix check: `IsSameOrDescendantDirectory` at lines 246–253
  - Documented invariant: `DESIGN.md` line 60; fix called out in `CHANGELOG.md` 3.0.2 (self-copy rejection)
  - Unit test only covers same lexical namespace: `tests/.../AlphaFS_Directory.Copy_DestinationRootEdgeCases.cs` (`DestinationInsideSource_Throws...`)

```246:253:src/AlphaFS/Filesystem/Directory Class/Directory Core Methods/Directory.CopyMoveCore.cs
      private static bool IsSameOrDescendantDirectory(string sourcePath, string destinationPath)
      {
         var sourcePathWithoutSeparator = sourcePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
         var destinationPathWithoutSeparator = destinationPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

         return destinationPathWithoutSeparator.Equals(sourcePathWithoutSeparator, StringComparison.OrdinalIgnoreCase) ||
                destinationPathWithoutSeparator.StartsWith(sourcePathWithoutSeparator + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
      }
```

- **Attack path:**
  1. Privileged or automated consumer calls `Directory.Copy` / cross-volume `Directory.Move` (`CopyAllowed` → emulated copy) with attacker-influenced source and destination strings.
  2. Attacker supplies the same directory under two path namespaces so the string prefix check fails, for example:
     - Source `C:\data` → `\\?\C:\data`
     - Dest `\\127.0.0.1\C$\data\out` → `\\?\UNC\127.0.0.1\C$\data\out`
     - Or Volume GUID vs drive letter; or 8.3 short name vs long name for the same directory
     - Or destination path that walks through a junction whose target is under the source (lexical dest outside source, final path inside)
  3. `CopyMoveDirectoryCore` enumerates the source while writing into a path that resolves inside the source → nested self-copy / unbounded growth (and for emulated move, later deletes the source after corrupting the tree).

- **Impact:** Denial of service (disk fill), tree corruption, incomplete/emulated-move data loss. Bypasses the explicit 3.0.2 safety fix and the DESIGN.md invariant.

- **Remediation:** Before the guard, canonicalize both paths to a comparable form (e.g. open with backup semantics, `GetFinalPathNameByHandle`, normalize Volume GUID + expand 8.3 via `GetLongFrom83ShortPath`, and/or equate UNC admin-share roots with drive letters). Compare volume identity + final path prefix, not raw strings. Apply the same check to emulated and non-emulated folder moves.

- **Why existing controls fail:** `GetFullPathName` (via `GetExtendedLengthPathCore`) does not unify drive letter, `\\?\Volume{GUID}\`, UNC admin shares, or 8.3 aliases, and does not resolve intermediate reparse points. The guard never calls `Volume.IsSameVolume` / final-path APIs that the codebase already has for move volume detection.

---

### 2. Recursive `Directory.Delete` remains path-based after enumeration (TOCTOU reparse swap)

- **Severity:** Medium
- **Location:** `src/AlphaFS/Filesystem/Directory Class/Directory Core Methods/Directory.DeleteDirectoryCore.cs`
- **Evidence:**
  - Non-follow policy at enumeration: lines 69–82 (`RecursionFilter = entry => !entry.IsReparsePoint`)
  - File deletes use path strings immediately: lines 93–96 → `File.DeleteFileCore`
  - Directory deletes later via path: lines 100–101 → `DeleteDirectoryNative` → `RemoveDirectory(pathLp)`
  - `DeleteFile` / `RemoveDirectory` resolve intermediate reparse points; they are not opened with `FILE_FLAG_OPEN_REPARSE_POINT` per path component
  - 3.0.1 changelog claimed recursive delete no longer modifies targets outside the specified tree via junctions/symlinks present at enumeration time

- **Attack path:**
  1. Elevated consumer recursively deletes a user-writable directory with AlphaFS.
  2. During enumeration of a still-normal subdirectory `A`, the attacker replaces `A` with a junction/symlink to a sensitive directory (e.g. under `C:\Windows` or another tenant path).
  3. Subsequent `DeleteFile`/`RemoveDirectory` on previously yielded child paths (`...\A\child`) now operate on the junction target.

- **Impact:** Confused-deputy deletion of files outside the intended tree (integrity loss; privilege impact when the consumer is elevated).

- **Remediation:** Prefer handle-relative deletion (open each directory with `FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT`, delete children via the directory handle / `NtSetInformationFile` disposition, or re-check reparse tag + file ID immediately before each destructive call and abort on change). Document residual TOCTOU if not fixed.

- **Why existing controls fail:** Controls only skip *recursing into* entries that are reparse points at find time. They do not bind the delete to a stable file ID or prevent an intermediate path component from becoming a reparse point between find and delete.

## Candidates investigated and rejected

| Candidate | Why rejected |
| --- | --- |
| Default directory copy follows directory symlinks unless `CopyOptions.CopySymbolicLink` | Intentional: junctions always re-created as links; symlinks follow `CopyFileEx`/`CopySymbolicLink` semantics. DESIGN.md only promises non-traversal for reparse points *selected for link copying*. Not a bypass of documented safety. |
| Junction/symlink following in Encrypt/Compress/DeleteEmptySubdirectories/GetSize | 3.0.1 paths force `SkipReparsePoints` (or equivalent). Static planting without a race does not redirect those operations. |
| `GetAccessControl` / `SetAccessControl` following links | Path-based Win32 security APIs; no extra AlphaFS “safe tree” claim. PrivilegeEnabler for SACL is documented/required behavior. |
| `Path.GetFullPath` early-return for `\\?\GlobalRoot\`, `\\?\Volume`, `\??\` without stripping `..` | Real canonicalization gap for niche prefixes, but documented special-path handling; complete high-impact chain against AlphaFS’s own invariants was weaker than findings above. Consumer allowlist misuse is possible but not an AlphaFS safety-invariant bypass by itself. |
| `..`, trailing dots/spaces, `\\?\` on normal drive paths | Relative paths go through `GetFullPathName` + `NormalizePath`; drive-letter self-copy is blocked by the lexical test. Not a bypass once both paths share the same normalized drive form. |
| Device names (`CON`/`NUL`/`\\.\PhysicalDrive`) | CreateFile device access is explicit Win32 surface; “API opens the path you pass” — out of scope per review rules. |
| `MoveOptions.DelayUntilReboot` → `PendingFileRenameOperations` | Documented admin/LocalSystem Win32 feature exposed as `MoveOptions`. Validation rejects UNC + `CopyAllowed`. Dangerous if a consumer forwards untrusted paths *and* the flag, but that is intentional API capability, not a bypass of a library safety control. `DirectoryInfo.MoveTo` even rejects the flag. |
| Alternate data streams unexpected access | Colon in paths often hits `CheckSupportedPathFormat`; ADS is reached via dedicated enumerate/backup APIs. No unexpected stream smuggling into ordinary File APIs beyond explicit ADS features. |
| Same-volume `Directory.Move` into descendant without `IsSameOrDescendantDirectory` | Check is gated on `cma.IsCopy` only; native `MoveFile` into a descendant typically fails without recursive self-copy. Emulated moves set `IsCopy` and *are* covered (subject to finding 1). |
| Copy/move TOCTOU between guard and operate | Classic race exists in principle; primary validated issue is the incomplete guard itself (finding 1). |

## Verdict

Two validated findings at Medium or higher: one High (incomplete destination-under-source guard), one Medium (delete TOCTOU after the 3.0.1 reparse fix). No Critical remote RCE-style issue; impact is local confused-deputy / safety-invariant failure for consumers that pass attacker-influenced paths into recursive copy/move/delete.
