# Maintainer Notes

This guide covers building, testing, release procedures, and durable implementation findings. Start with the [README](../README.md#library-and-api-usage) for library usage and an API preview.

Keep unfinished work in [Email.ToolKit.Tasks.md](Email.ToolKit.Tasks.md), significant release changes in [ReleaseNotes.txt](ReleaseNotes.txt), and architectural contracts in [OutlookLifecycle.md](OutlookLifecycle.md). The [historical API audit](ApiCompatibilityAudit.md) and [duplicate comparison reference](Email.ToolKit.DuplicateComparison.md) retain compatibility evidence and client-visible property mappings. Completed task lists and routine validation diaries are not maintained.

Use a Visual Studio Developer Command Prompt with a compatible .NET SDK, Git, GitHub CLI, and `zip.exe` available. Visual Studio MSBuild is required for the Outlook COM references. Commands below run from the repository root.

## Contributing

### Contribution workflow

1. Fork the Project
2. Create your Bug / Feature Branch (`git checkout -b feature/amazing-feature`)
3. Commit your Changes (`git commit -m 'Add some amazing feature'`)
4. Push to the Branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Coding style

Please match the current coding style.  Most notably:

1. One operation per line
2. Use complete English words in variable and method names
3. Attempt to declare variable and method names in a self-documenting manner
4. Add unit tests

## Local builds

`DevelopmentTools\build.cmd` rebuilds Debug. For cleanup, run `DevelopmentTools\clean.cmd`. The `release` and `publish` options distribute artifacts; see [Preparing a release](#preparing-a-release).

`DevelopmentTools\test.cmd` builds Debug with Visual Studio MSBuild and runs all four solution test projects through their NUnit adapters, including real-Outlook integration tests. Those tests require a working Outlook installation and profile. NUnit `[Explicit]` and `[Ignore]` exclusions still apply. TRX results are under `artifacts/TestResults`; extra `dotnet test` arguments are forwarded, for example `test.cmd --list-tests` or `test.cmd --filter "TestCategory=Compatibility"`.

### Test organization

Shared helpers/fakes are in `SourceCode/ToolKit.Tests/Support`, with embedded fixtures in `Support/Resources`. Explicit logical resource names preserve fixture lookups. Keep methods alphabetical and execution independent; retain STA and nonparallel requirements for real COM tests. Move files separately from content edits when preserving clear Git rename history.

The application and library embed Release PDB symbols. Library XML documentation remains in the NuGet package for consumer IntelliSense. Artifacts are under `artifacts/<version>/`; fresh `publish-*` staging directories prevent stale files entering executable archives.

## Versioning

Starting with `1.14.0`, releases use `major.minor.patch` numbering. The third component is a patch version, replacing the previous build-number sequence. Do not increment the package version merely because a build ran.

- Compatible fixes increment the patch: `1.14.0` to `1.14.1`.
- Compatible features increment the minor: `1.14.x` to `1.15.0`.
- Intentional breaking public API changes increment the major version.

Documentation may name an upcoming release before the project versions are changed. Reserve the actual version update for one atomic commit near the end of release preparation, before building the final release artifacts.

Version values for `1.14.1` in the application and library:

```xml
<AssemblyVersion>1.0.0.0</AssemblyVersion>
<FileVersion>1.14.1.0</FileVersion>
<Version>1.14.1</Version>
```

Keep `AssemblyVersion` stable for compatible 1.x releases. Update `Version` and `FileVersion` together in both application and library project files, using the selected release version.

## Public API deprecation review

Version 1.14.1 restored two migration overloads with warning-only obsolete annotations and corrected existing replacement guidance. All annotated methods remain callable; the table below describes migration paths.

| Legacy API | Migration guidance |
| --- | --- |
| MapiItem.DeleteDuplicate | Use OutlookItem.DeleteDuplicate. |
| MapiItem.DeleteItem | Use OutlookItem.Delete(item). |
| MapiItem.GetItemHash | Use new OutlookItem(item).Hash. |
| MapiItem.GetItemHashAsync | Use OutlookItem.GetHashAsync(item). |
| MapiItem.GetItemSynopses | Use new OutlookItem(item).Synopses. |
| MapiItem.GetPath | Use OutlookItem.GetPath(item). |
| MapiItem.MoveItem / MoveItemAsync | Use OutlookItem.Move / MoveAsync with item and destination. |
| MapiItem.RemoveMimeOleVersion | Use OutlookItem.RemoveMimeOleVersion(header). |
| OutlookFolder.GetMailItemSynopses | Use new OutlookItem(mailItem).Synopses. |
| OutlookStore.GetItemSynopses | Get the item with OutlookSession.GetItemFromId, then read OutlookItem.Synopses. |
| OutlookStore.GetFolderFromID / GetFolderFromId | Use OutlookSession.GetFolderFromId(entryId, store.StoreID); the result is an OutlookFolder wrapper. |
| Static OutlookFolder.GetSendersCount with path/dictionary | Use GetSendersCount(folder) on an OutlookFolder instance. The legacy signature has no int parameter. |
| OutlookFolder.RemoveFolder with path | Use RemoveFolder(subFolder, subFolderIndex, force). |
| OutlookFolder.MergeFolders / MoveFolderContents / RemoveDuplicates with path | Use the corresponding instance overload without path. Preserve dryRun where applicable. |
| List-based OutlookItem.GetDateTimesBytes / GetDateTimesText | Use the ReadOnlyCollection overloads. Review null handling during conversion. |
| Migrate.EmlFileToPst(filePath, pstPath) | Supply closeStore explicitly; false preserves the old attached-store behavior. |
| Migrate.EmlToPst(path, pstPath, adjust) | Supply closeStore explicitly; false preserves the old attached-store behavior. |

The restored two-argument `OutlookFolder.AddFolder` remains supported without an obsolete annotation. A useful default-preserving convenience overload need not be deprecated merely because it also restores binary compatibility.

New migration annotations produce CS0618 on recompilation. Consumers using warnings as errors may need to update their calls. Precompiled callers can still bind to the restored signatures; the attributes do not detach stores or otherwise change migration behavior. Compatibility tests deliberately exercise the warning-producing calls, with CS0618 disabled only at those call sites.

Use this inventory to inform version 2 design; removal is not automatic. Keep the legacy APIs in 1.x, verify migration paths before adding further warnings, and assess raw COM return types and null behavior explicitly. `Migrate`, `OutlookAccount`, and account-based constructors were not broadly deprecated because their full replacement paths remain design work.

## Preparing a release

1. Select scope and review dependencies. Publish required upstream/submodule packages first; commit and push submodule changes before their parent references.
2. Compare public APIs against the last published NuGet package and any deliberately supported older baseline. A GitHub tag is not evidence of NuGet publication. Add tests before changing legacy contracts.
3. Validate every Release target, relevant NUnit tests, consumer compilation/binary binding, and packaged MIME conversion. Extract the executable ZIP separately and smoke-test help, `--timeout` / `-t`, failed connections, and normal Outlook operations. See [Outlook verification](#outlook-verification).
4. Finalize `Documentation/ReleaseNotes.txt`. Update application/library `Version` and `FileVersion` together in one commit near the end; retain `AssemblyVersion` for compatible releases.
5. Let the maintainer commit, merge to `main`, and push. Both distribution commands require `main` and a clean working tree, including submodules.
6. Run `DevelopmentTools\build.cmd release`, then `DevelopmentTools\build.cmd publish` when distributing both artifacts. Verify the GitHub release/tag and NuGet workflow result.
7. Remove finished backlog entries after retaining significant release changes and durable decisions.

The commands read project versions automatically; an optional second argument checks an expected version. The first command creates the version's GitHub release and tag if absent; the second reuses them. Existing published tags must not be moved. Tag/commit verification hardening remains tracked in the backlog.

### Explicit publication

`release` builds a self-contained, single-file Windows x64 executable, creates a fresh ZIP using `zip.exe`, and uploads it to a published GitHub release. `publish` builds/packs all library Release targets locally, uploads the package to the matching release, and dispatches `build.yml` from `main`. Either command can run independently; retries replace only its own asset. An existing draft is published automatically.

Local builds handle the Outlook COM references. The Ubuntu workflow downloads the prebuilt package and uses `NuGet/login@v1` for a short-lived publishing key; it does not build Office-dependent code. NuGet publication requires explicit workflow dispatch. Pushes, tags, and GitHub release events do not trigger it.

Configure the NuGet trusted-publishing policy for owner `jamesjohnmcguire`, repository `DigitalZenWorks.Email.ToolKit`, workflow `build.yml`, with no environment. GitHub CLI uses `GH_TOKEN`; the workflow uses GitHub's automatic token to download the package.

A successful `build.cmd publish` means dispatch was accepted, not that NuGet publication succeeded. Follow the printed run link or watch completion:

```shell
gh run watch --repo jamesjohnmcguire/DigitalZenWorks.Email.ToolKit --exit-status
```

For failures, inspect `gh run view <run-id> --log-failed` before retrying. After a successful push, NuGet indexing can delay search visibility.

## Outlook verification

The [lifecycle design](OutlookLifecycle.md) defines connection, ownership, and timeout contracts. Unit tests exercise coordination through established abstractions; tests activating Outlook or creating PSTs remain integration tests.

### Temporary probe boundary

Do not return or cache the temporary probe's COM application: its STA exits. The probe releases its own reference on its worker; the real application/session is acquired on the caller's STA. This deliberate second acquisition avoids the apartment-lifetime defect. The probe neither calls `Quit` nor forcibly releases a shared RCW to zero. See [the lifecycle explanation](OutlookLifecycle.md#keep-the-temporary-probe-separate-from-the-real-connection).

Probe failures must reach `ProbeAttempt.Run` for capture/logging, including after timeout. A catch that merely logs and returns would falsely report success. Worker cleanup must not signal an already-disposed wait handle.

### Manual regression checks

| Initial state | Checks |
| --- | --- |
| Healthy Outlook running | Connect succeeds; repeated calls reuse the session; Disconnect leaves Outlook usable. |
| Outlook closed with a configured profile | Connect activates Outlook and obtains a reusable session. |
| Unresponsive Outlook | Probe wait returns false; no service connection/session; retries do not accumulate blocked probes. |
| Recovery after timeout | Let the outstanding worker finish, then retry successfully in the same host. |

Use the same service and caller STA within each scenario. Ordinary integration setup starts Outlook, so it cannot establish initially closed/unresponsive states. Capture versions/bitness, profile, reproduction steps, timeout, elapsed time, and logged stage when investigating a failure; retain useful evidence with the issue rather than in a permanent validation diary.

Deliberate hangs belong in a disposable interactive environment with a watchdog around the test host. Watchdog expiry is an aborted test, not proof of production cancellation. A suspended process does not establish reproduction of the original failure. Healthy integration results alone do not establish hung-Outlook behavior.

### Outlook/PST resource retention

Repeated full integration runs have produced `AddStore` failures (`0x80004005`, `0x80040119`, `0x80040111`), including when Visual Studio and the command line used the same assemblies and ordinary user context. Restarting Outlook and/or closing stale temporary stores restored passing runs, but the underlying resource-retention cause remains unconfirmed.

PST detachment does not guarantee immediate file unlock; MSG files can also remain locked. Test contexts use GUID-unique PST filenames and preserve backing files if detachment fails. Best-effort deletion can still report retained locks. Do not add fixed sleeps, delete still-attached backing files, or automatically terminate the user's Outlook as a presumed fix.

### Optional VM testing

An on-demand Windows VM with a dedicated profile, test data, snapshots, and a signed-in interactive desktop can isolate failure scenarios. It supplements deterministic unit tests and still needs explicit reproduction steps. See Microsoft's [desktop test guidance](https://learn.microsoft.com/en-us/azure/devops/pipelines/test/ui-testing-considerations?view=azure-devops) and [Office automation limitations](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa).

## Dependency regression checks

MsgKit 3.0.6 requires MimeKit 4.x; keep MimeKit at 4.18.2 until an upstream update supports the split assemblies introduced in MimeKit 5. Explicit OpenMcdf 3.4.0 and System.Security.Cryptography.Pkcs 11.0.0 references provide security minimums (the latter outside the .NET Framework target). Confirm those minimums in generated package metadata. Remove interim references only when upstream dependencies retain the fixes.

`MsgKitConversionTests` exercises the production adapter without Outlook, checking attachment bytes, names, IDs, HTML references, read-receipt fields, nested/attached messages, and unchanged source files. `MsgKitAttachmentTests` retains a valid upstream inline-attachment control. Check retained content, not merely the absence of exceptions.

## MsgKit MIME Compatibility

### Valid MIME and optional identifiers

RFC 2045 section 7 makes Content-ID generally optional. RFC 2183 section 2.1 uses `inline` to request automatic display; it does not require an HTML `cid:` reference. RFC 2298 includes a read-receipt example whose notification part has no Content-ID. Such parts are not inherently malformed.

An unmatched HTML `cid:` reference is a separate issue. RFC 2110 also allows related parts to be identified by Content-Location. The original archived message that triggered the constructor exception is still unavailable.

### Upstream behavior and fork removal

Synthetic conversion cases compared the former fork with public MsgKit 3.0.4 and 3.0.6. The converter preserved the tested loose attachments despite the direct constructor rejecting inline attachments without an ID. The fork was therefore unnecessary for those EML conversion paths. A separate read-receipt data-loss defect required a local adapter.

| Case | Unmodified converter result |
| --- | --- |
| Missing, blank, or unusable Content-ID | Attachment bytes retained |
| Missing filename | Bytes retained; fallback name supplied |
| Matching HTML cid | Reference, ID, and bytes retained |
| Unmatched HTML cid | Reference retained; no target ID invented |
| Relative Content-Location | Bytes/name/HTML retained; location tag absent |
| Read receipt with or without inline disposition | Notification data empty |

The notification tests verified that recipient and disposition fields were present before conversion. The specialized MsgKit notification branch set a filename but never wrote those fields into its attachment stream. This is a data-preservation defect, separate from constructor validation.

The earlier constructor probe demonstrated that upstream rejects a direct `Attachments.Add` call with `isInline: true` and a missing ID. The EML converter avoids that condition for the examples above. Production regressions therefore exercise conversion, rather than require the upstream constructor to accept values it intentionally rejects.

### Email.ToolKit workaround

The library now uses public MsgKit 3.0.6 and MimeKit 4.18.2. `EmlConverter` inspects the MIME body and replaces disposition-notification parts only in an internal conversion copy. Their serialized fields become ordinary attachments, defaulting to `read-receipt.txt`. Existing names and Content-IDs are retained. No identifiers are generated.

The archived EML is never rewritten. Messages without notifications are passed to MsgKit from the original input bytes. Attached messages remain attached EML files; their internal MIME bodies are not normalized.

This preserves notification information as readable attachment data. It does not create a native Outlook read-receipt object or restore unknown HTML relationships. MIME inspection adds a parse to conversion; the inspected message is disposed before MsgKit parses the selected input.

### Regression scope and upstream follow-up

Permanent production-adapter tests cover the synthetic cases above, including notification fields, existing metadata, and source preservation. Real Outlook import/rendering and historical samples still need separate verification. No production build references the fork.

The [unsent issue draft](#unsent-upstream-issue-draft) records the read-receipt defect. An upstream response is not required for the local workaround; retire it only after a fixed package passes regression tests.

### References

- [MIME identifiers](https://www.rfc-editor.org/rfc/rfc2045.html#section-7)
- [Inline disposition](https://www.rfc-editor.org/rfc/rfc2183.html#section-2.1)
- [Read-receipt example](https://www.rfc-editor.org/rfc/rfc2298.html#section-9)
- [Content-Location](https://www.rfc-editor.org/rfc/rfc2110.html)
- [Upstream package](https://www.nuget.org/packages/MsgKit/3.0.6)

### Unsent upstream issue draft

Proposed title: ConvertEmlToMsg discards disposition-notification fields

This is a local issue draft. It has not been submitted upstream.

#### Problem

Converting a multipart/report containing a message/disposition-notification part creates an MSG attachment with an empty data stream. The input includes Final-Recipient, Original-Message-ID, and Disposition fields. The human-readable message body converts successfully, but the machine-readable fields are lost.

This reproduces with public MsgKit 3.0.6. It also reproduced with 3.0.4 and with the local attachment-tolerance fork. No Outlook installation or COM activation is needed to reproduce the conversion defect.

#### Input and reproduction

Use [the synthetic EML](Examples/ReadReceiptWithoutContentId.eml). Its missing Content-ID is intentional: RFC 2298 section 9 includes the same general MIME structure without a Content-ID on the notification part.

Create a console project targeting .NET 10 and reference MsgKit 3.0.6, MimeKit 4.18.2, OpenMcdf 3.4.0, and System.Security.Cryptography.Pkcs 11.0.0. Copy the EML into the working directory and run:

```csharp
using System;
using MsgKit;
using OpenMcdf;

Converter.ConvertEmlToMsg(
    "ReadReceiptWithoutContentId.eml", "receipt.msg");

using RootStorage root = RootStorage.OpenRead("receipt.msg");
Storage attachment = root.OpenStorage("__attach_version1.0_#00000000");
using CfbStream data = attachment.OpenStream(
    PropertyTags.PR_ATTACH_DATA_BIN.Name);
Console.WriteLine(data.Length);
```

Start with no existing `receipt.msg`; the converter creates a new file. Expected: a nonempty attachment containing the notification fields. Actual: the printed attachment length is zero.

#### Apparent cause

In `Converter.ConvertEmlToMsg`, this branch sets the filename and exits:

```csharp
case MessageDispositionNotification notification:
    fileName = notification.FileName;
    break;
```

Unlike the delivery-status branch, it does not write the notification body to `attachmentStream`. One possible upstream correction is to serialize its content before exiting, then verify the decoded MSG attachment fields.

The regression should assert retained recipient, original-message ID, and disposition values, both with and without an inline disposition header. Checking only that conversion does not throw would miss the data loss.

#### Scope

This report concerns preservation of parsed notification data. It does not request a change to attachment-constructor validation, or assume every notification is referenced by an HTML cid URL.

Email.ToolKit has a local workaround that represents these fields as an ordinary attachment in a conversion copy. An upstream correction would let that workaround be retired once a fixed package is available and tested.

- [RFC 2298 example](https://www.rfc-editor.org/rfc/rfc2298.html#section-9)
- [MsgKit package](https://www.nuget.org/packages/MsgKit/3.0.6)
- [Current converter source](https://github.com/Sicos1977/MsgKit/blob/master/MsgKit/Converter.cs)

# Additional Notes

## Building
	For one-off builds, use:
		MSBuild.exe -p:Configuration=Release;DebugType=embedded;IncludeNativeLibrariesForSelfExtract=true;PublishDir=%CD%\Release\win-x64;PublishDocumentationFiles=false;PublishReadyToRun=true;PublishSingleFile=true;RuntimeIdentifier=win-x64;SelfContained=true -restore -t:Publish ToolKit.Application
