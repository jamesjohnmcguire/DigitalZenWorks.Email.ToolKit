# NuGet API Compatibility Audit

Snapshot: 2026-11-07T02:14:05.467578+00:00 (UTC).

Package: `DigitalZenWorks.Email.ToolKit`, `1.4.22` through `1.14.0`.

## Findings and observed impact

Published assemblies contain historical API incompatibilities starting with
`1.4.22 -> 1.4.23`; they were not all introduced in 1.13 or 1.14.

The maintainer reports no API-breakage complaints and expects light/casual
usage. No affected consumer has been identified. This is a compatibility
inventory, not evidence of widespread failures or an emergency repair list.

- All 18 currently downloadable NuGet versions were inspected; all are listed.
- 74 packaged library assemblies were inspected without executing them.
- 17 adjacent-version and 17 historical-to-1.14.0 comparisons produced
  33 distinct package-pair runs; the final pair belongs to both sets.
- 11 of 17 adjacent transitions emitted compatibility diagnostics.
- 38 distinct type/member diagnostic identities occur across those
  transitions; 37 remain absent from 1.14.0. Counts include generated
  delegate methods, not just independent design changes or affected clients.
- All 18 within-package cross-framework API checks passed.
- All 74 assemblies retain the unsigned identity
  `DigitalZenWorks.Email.ToolKit.Library, Version=1.0.0.0`.
- No standalone parameter-name diagnostic was emitted for matching members.
  No comparison failed because of unresolved references or tool errors.
- A separate check found 24 framework-label mismatches in 12 older packages.
  No such mismatch was found in 1.14.0.

Only three signatures are missing from the most recent prior NuGet package,
**1.12.62**, when compared directly with **1.14.0**. Those are the most
focused candidates for a future compatibility patch.

## Method and evidence

Microsoft's `Microsoft.DotNet.ApiCompat.Tool` **10.0.403** compared the actual
NuGet `.nupkg` files, rather than rebuilt historical source. Public/protected
APIs were checked, internal APIs excluded, and additive APIs allowed.
Parameter-name checking was enabled. No suppressions were used.

Separate self-checks enabled strict API comparisons between compatible
framework assets within each package. A temporary metadata reader recorded
assembly identities, framework attributes, references, and method metadata
without activating Outlook or creating COM objects.

`CP0001` means a baseline type is absent. `CP0002` means an exact baseline
member signature is absent. The latter can reflect removal, renaming, or
return/parameter changes; the method name may still exist. An optional new
parameter does not retain an old binary signature.

The [NuGet package-content API][nuget-api] returns listed and unlisted
currently downloadable versions. This snapshot found no unlisted versions.
Deleted or unavailable artifacts cannot be inspected from this inventory.
See [Microsoft's ApiCompat documentation][apicompat] for the tool modes.

[nuget-api]: https://learn.microsoft.com/en-us/nuget/api/package-base-address-resource
[apicompat]: https://learn.microsoft.com/en-us/dotnet/fundamentals/apicompat/global-tool

## Published-package inventory

Framework sets use NuGet folder names; `win` abbreviates `-windows7.0`.

- **A:** `net461`, `net5.0-win`, `net6.0-win`.
- **B:** `net472`, `netstandard2.0`, `net5.0-win`, `net6.0-win`, `net7.0-win`.
- **C:** B plus `net8.0-win`.
- **D:** C plus `net9.0-win`.
- **E:** D plus `net10.0-win`.

Set A contains the framework-label problem described below. Later sets have
matching base target-framework attributes.

| NuGet version | Published (UTC) | Assets | Embedded source commit |
| --- | --- | --- | --- |
| 1.4.22 | 2022-04-07 | A | Not recorded |
| 1.4.23 | 2022-04-08 | A | Not recorded |
| 1.4.24 | 2022-04-08 | A | Not recorded |
| 1.5.25 | 2022-04-11 | A | Not recorded |
| 1.5.26 | 2022-04-25 | A | Not recorded |
| 1.5.27 | 2022-05-04 | A | Not recorded |
| 1.6.28 | 2022-05-20 | A | Not recorded |
| 1.6.29 | 2022-06-10 | A | Not recorded |
| 1.6.30 | 2022-07-02 | A | Not recorded |
| 1.6.31 | 2022-09-30 | A | Not recorded |
| 1.6.34 | 2022-11-11 | A | Not recorded |
| 1.7.42 | 2022-12-26 | A | Not recorded |
| 1.8.50 | 2023-10-20 | B | Not recorded |
| 1.8.54 | 2024-08-21 | C | `5e60271` |
| 1.9.56 | 2024-08-31 | C | `ef19ee2` |
| 1.10.58 | 2024-09-16 | C | `cc35c99` |
| 1.12.62 | 2025-03-10 | D | `5ea161e` |
| 1.14.0 | 2026-11-07 | E | `398bc4f` |

Five packages record Git commits in their manifests. For older packages,
Git tags are supporting evidence, not proof of the exact build commit.
Several prefixed/unprefixed tags point at different commits. Actual package
binaries are authoritative for this report.

Several GitHub-only versions, including 1.12.64 and 1.13.65, have no package
in this NuGet snapshot. Early source tags such as `1.4.22.0` correspond to
normalized NuGet package version `1.4.22`; the source tags are unchanged.

## Adjacent published-version results

| Transition | Distinct diagnostics | Status in 1.14.0 |
| --- | ---: | --- |
| 1.4.22 -> 1.4.23 | 2 | 2 still absent |
| 1.4.23 -> 1.4.24 | 0 | No diagnostic |
| 1.4.24 -> 1.5.25 | 1 | 1 still absent |
| 1.5.25 -> 1.5.26 | 9 | 9 still absent |
| 1.5.26 -> 1.5.27 | 2 | 2 still absent |
| 1.5.27 -> 1.6.28 | 3 | 3 still absent |
| 1.6.28 -> 1.6.29 | 1 | 1 restored; 0 still absent |
| 1.6.29 -> 1.6.30 | 0 | No diagnostic |
| 1.6.30 -> 1.6.31 | 0 | No diagnostic |
| 1.6.31 -> 1.6.34 | 0 | No diagnostic |
| 1.6.34 -> 1.7.42 | 9 | 9 still absent |
| 1.7.42 -> 1.8.50 | 0 | No diagnostic |
| 1.8.50 -> 1.8.54 | 1 | 1 still absent |
| 1.8.54 -> 1.9.56 | 6 | 6 still absent |
| 1.9.56 -> 1.10.58 | 0 | No diagnostic |
| 1.10.58 -> 1.12.62 | 1 | 1 still absent |
| 1.12.62 -> 1.14.0 | 3 | 3 still absent |

A clean row means no diagnostic under the selected rules; it does not prove
identical behavior or complete source compatibility. Packaged method
metadata confirms that `Migrate.EmlToPst(string, string)` disappeared in
1.6.29 and returned in 1.7.42. It is present in 1.14.0.

## Most recent NuGet transition: 1.12.62 to 1.14.0

These three exact signatures are absent in 1.14.0:

```csharp
Outlook.MailItem Migrate.EmlFileToPst(string filePath, string pstPath)
bool Migrate.EmlToPst(string path, string pstPath, bool adjust)
Outlook.MAPIFolder OutlookFolder.AddFolder(
    Outlook.MAPIFolder parentFolder, string folderName)
```

The migration methods gained a required `closeStore` parameter. Both source
calls to the old overloads and precompiled callers can be affected.
`AddFolder` gained optional `normalizeFolderName = false`: old two-argument
source calls can recompile, but precompiled callers request the old member.

The close-store changes are already in GitHub's 1.12.64 and v1.13.65 source
tags. The AddFolder change is absent from v1.13.65 and present in 1.14.0.
Relevant source commits include `4e663f0` (close-store option) and `14f687a`
(folder-name normalization). Their first NuGet publication is 1.14.0.

Forwarding overloads are a small additive maintenance option. Tests should
verify the exact old signatures and their earlier defaults/behavior. No
repair has been made by this audit.

## Historical signature details

Signatures below omit the ToolKit, collections, task, and basic system
namespaces. `Outlook` means `Microsoft.Office.Interop.Outlook`. A removed
type includes its contract; CP0001 does not count every member separately.

### 1.4.22 -> 1.4.23

The store-based count changed from uint to int. The path/folder overload was
replaced by a different count-oriented signature.

```text
bool OutlookStore.RemoveEmptyFolders(string, Outlook.MAPIFolder)
uint OutlookStore.RemoveEmptyFolders(Outlook.Store)
```

### 1.4.24 -> 1.5.25

The public parameterless folder constructor disappeared when account-based
construction was introduced.

```text
OutlookFolder()
```

### 1.5.25 -> 1.5.26

Refactoring removed a public mapper class and several traversal/cleanup
signatures. Some names remain with changed parameters or return types.

```text
Type: MessageMapper
int[] OutlookFolder.RemoveDuplicates(Outlook.MAPIFolder, bool, bool)
int[] OutlookFolder.RemoveDuplicatesFromSubFolders(string, Outlook.MAPIFolder,
    bool)
uint OutlookStore.MergeFolders(Outlook.Store)
void OutlookAccount.MergeFolders()
void OutlookAccount.RemoveEmptyFolders()
void OutlookFolder.MergeFolders(string, Outlook.MAPIFolder)
void OutlookStore.MergeFolders(string)
void OutlookStore.RemoveEmptyFolders(string)
```

### 1.5.26 -> 1.5.27

The public FolderMapper helper class and TransferFolderCallBackType delegate
disappeared.

```text
Type: FolderMapper
Type: TransferFolderCallBackType
```

### 1.5.27 -> 1.6.28

CreaterFolderPath was renamed to CreateFolderPath. GetSubFolder gained a bool
parameter. A cleanup traversal signature disappeared.

```text
Outlook.MAPIFolder OutlookFolder.CreaterFolderPath(Outlook.Store, string)
Outlook.MAPIFolder OutlookFolder.GetSubFolder(Outlook.MAPIFolder, string)
int OutlookStore.RemoveEmptyFolders(string, Outlook.MAPIFolder, int)
```

### 1.6.28 -> 1.6.29

The two-argument migration overload disappeared temporarily; it was restored
in 1.7.42.

```text
bool Migrate.EmlToPst(string, string)
```

### 1.6.34 -> 1.7.42

Traversal APIs and delegates were reorganized. FolderAction changed its
callback signature, affecting generated Invoke/BeginInvoke/EndInvoke members.

```text
Type: FolderActionConditional
IDictionary<string, IList<string>> OutlookFolder.GetItemHashes(string,
    Outlook.MAPIFolder, IDictionary<string, IList<string>>)
IAsyncResult FolderAction.BeginInvoke(string, Outlook.MAPIFolder,
    AsyncCallback, object)
int OutlookFolder.RecurseFolders(string, Outlook.MAPIFolder, bool,
    FolderActionConditional)
int OutlookFolder.RemoveEmptyFolders(string, Outlook.MAPIFolder, bool)
int[] OutlookFolder.RemoveDuplicates(string, Outlook.MAPIFolder, bool)
string MapiItem.GetItemHash(string, Outlook.MailItem)
void FolderAction.EndInvoke(IAsyncResult)
void FolderAction.Invoke(string, Outlook.MAPIFolder)
```

### 1.8.50 -> 1.8.54

AddMsgFile changed its return type from void to Outlook.MailItem. Source
statement calls can recompile; the binary signature changed.

```text
void OutlookFolder.AddMsgFile(Outlook.MAPIFolder, string)
```

### 1.8.54 -> 1.9.56

Several MapiItem methods widened MailItem parameters to object.
Moveitem/MoveitemAsync were renamed to MoveItem/MoveItemAsync. Parameter
widening can retain source calls while losing the old binary signature.

```text
Task MapiItem.MoveitemAsync(object, Outlook.MAPIFolder)
Task<string> MapiItem.GetItemHashAsync(Outlook.MailItem)
string MapiItem.GetItemHash(Outlook.MailItem)
string MapiItem.GetItemSynopses(Outlook.MailItem)
string MapiItem.GetPath(Outlook.MailItem)
void MapiItem.Moveitem(object, Outlook.MAPIFolder)
```

### 1.10.58 -> 1.12.62

The IList<DateTime> overload disappeared. Current List<DateTime> and
ReadOnlyCollection<DateTime> overloads do not retain that exact signature.

```text
byte[] OutlookItem.GetDateTimesBytes(IList<DateTime>)
```

### 1.12.62 -> 1.14.0

See the three recent overload gaps above.

```text
Outlook.MAPIFolder OutlookFolder.AddFolder(Outlook.MAPIFolder, string)
Outlook.MailItem Migrate.EmlFileToPst(string, string)
bool Migrate.EmlToPst(string, string, bool)
```

## Historical framework-packaging mismatch

All 12 published versions from **1.4.22 through 1.7.42** contain .NET 6
assemblies in both `lib/net461` and `lib/net5.0-windows7.0`: 24 mislabelled
assets. Their `TargetFrameworkAttribute` is `.NETCoreApp,Version=v6.0`, and
they reference .NET 6 framework assemblies.

In 1.4.22, all three framework folders contain the identical DLL, SHA-256:

```text
3c8b7ca33e28c88de5911d05dbec64d94a22b8398306166ad37e8b3355155c96
```

This packaging/targeting issue is separate from member changes. Those old
folders cannot be relied upon as evidence of genuine .NET Framework 4.6.1
or .NET 5 support. Historical runtime loading was not tested, and no
consumer failure has been reported.

**1.8.50 is the first inspected published version with matching framework
metadata in every folder.** Later inspected packages, including 1.14.0,
also match. No current-package change is needed for this historical finding.

ApiCompat's within-package checks still pass for old packages: they compare
API surfaces, not the accuracy of framework-folder labels. The independent
metadata inspection established this finding.

## Suggested disposition

With no reported failures and an expected small casual-client base, these
metadata findings alone do not justify an emergency rollback or recreating
every historical API.

1. Consider the three 1.12.62 overloads for the next convenient compatibility
   patch, with tests preserving the old signatures and defaults/behavior.
2. Document older removals. Restore an inexpensive alias/overload when useful
   or when an actual consumer needs it; do not expand the API automatically.
3. Do not blindly restore old return types or the old FolderAction signature.
   C# cannot overload solely by return type, and reversing those contracts
   could break later-version consumers. These need a deliberate decision.
4. Avoid introducing new public COM entry points solely to recreate old APIs.
5. Before future releases, compare with the last published NuGet package and
   any older versions the team deliberately supports. A GitHub tag alone is
   not the NuGet publication baseline.

No APIs, project versions, packages, releases, or workflows were changed by
this audit. Unrelated working-tree changes were preserved.

## Reproduction and retained artifacts

Downloads, checksums, metadata, commands, and logs are under ignored
`artifacts/compatibility-audit/`. No build dependency or CI changes were made.

Example package comparison:

```shell
artifacts\compatibility-audit\tools\apicompat.exe package artifacts\compatibility-audit\packages\1.14.0\digitalzenworks.email.toolkit.1.14.0.nupkg --baseline-package artifacts\compatibility-audit\packages\1.12.62\digitalzenworks.email.toolkit.1.12.62.nupkg --run-api-compat --enable-rule-cannot-change-parameter-name
```

- `inventory.json`: URLs, publication data, dependencies, source commits,
  and archive SHA-256 values.
- `assembly-metadata.json`: identities, frameworks, references, type names,
  and method metadata.
- `comparison-results.json`: all 33 commands, exit codes, and diagnostics.
- `self-results.json`: all 18 cross-framework consistency checks.
- `framework-mismatches.json`: the 24 mislabelled older assets.
- `source-map.json`: embedded source evidence and candidate Git tags.
- `results/*.log`: complete unsuppressed tool output.

Historical comparison exit code 1 represents the reported incompatibilities.
All failures had CP0001/CP0002 diagnostics; none was an unexplained tool
failure. All 18 within-package self-checks returned 0.

## Limits

This is an assembly-level API audit, not a behavioral, dependency-security,
or COM runtime test suite. It does not prove identical exception/null
handling, optional-default semantics, serialization, Outlook lifetime or
threading behavior, or every C# overload-resolution outcome. Attribute
matching was not enabled; nullability/attribute compatibility is not
exhaustively assessed. Unavailable/deleted package versions are outside the
snapshot. Older tag provenance is qualified when the package records no
source commit.
