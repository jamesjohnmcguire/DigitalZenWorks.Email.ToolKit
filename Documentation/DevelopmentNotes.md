# Maintainer Notes

This document is for contributors maintaining the repository: building,
testing, diagnosing implementation issues, and preparing releases. For
library usage and an API preview, start with the [README][api-overview] and
[Outlook lifecycle](OutlookLifecycle.md).

The filename remains `DevelopmentNotes.md` to preserve existing links.

[api-overview]: ../README.md#library-and-api-usage

Run commands from a Visual Studio Developer Command Prompt with .NET 10,
Git, and `zip.exe` available. The build uses Visual Studio MSBuild because the
Outlook COM references require it. Run commands from the repository root;
the script also works from another directory and restores that directory.

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

```shell
DevelopmentTools\build.cmd
DevelopmentTools\build.cmd release
```

The default command rebuilds Debug. `release` publishes a self-contained,
single-file Windows x64 application, creates a fresh ZIP with `zip.exe`,
and uploads only that ZIP to a GitHub draft release. `publish` builds and
packs all Release library targets, then publishes the NuGet package through
GitHub Actions. These commands are independent; neither requires the other
artifact. Run `DevelopmentTools\clean.cmd` directly for cleanup.

The application and library embed PDB symbols in their Release assemblies.
The library XML documentation remains in the NuGet package for consumer
IntelliSense; it is not required to run the application.

Artifacts are under `artifacts/<version>/`. The ZIP and package names include
the project version. Each executable build uses a new `publish-*` staging
directory so old publish files cannot enter the next archive. The `.commit`
files record the source commit of each prepared artifact. Local builds may
include uncommitted changes; release rehearsals are not final artifacts.

## Versioning

Starting with `1.14.0`, releases use `major.minor.patch` numbering. The third
component is a patch version, replacing the previous build-number sequence.
Do not increment the package version merely because a build ran.

- Compatible fixes increment the patch: `1.14.0` to `1.14.1`.
- Compatible features increment the minor: `1.14.x` to `1.15.0`.
- Intentional breaking public API changes increment the major version.

Documentation may name an upcoming release before the project versions are
changed. Reserve the actual version update for one atomic commit near the
end of release preparation, before building the final release artifacts.

Target values for `1.14.0`, to apply in that final version commit:

```xml
<AssemblyVersion>1.0.0.0</AssemblyVersion>
<FileVersion>1.14.0.0</FileVersion>
<Version>1.14.0</Version>
```

Keep `AssemblyVersion` stable for compatible 1.x releases. Update `Version`
and `FileVersion` together in both application and library project files,
for example `1.14.1` and `1.14.1.0` for a subsequent patch release.

## Preparing a release

Keep application and library `Version` values aligned and update their
four-segment `FileVersion` values. Preserve `AssemblyVersion` for compatible
releases. The script reads the versions automatically; an optional second
argument checks an expected version without changing it.

Review Release package references separately from Debug project references.
Build every Release target, run the relevant NUnit tests, and test the ZIP
and package. Verify `--timeout` / `-t`, help, connection failure, and normal
Outlook operations. Real Outlook tests require a working profile and test
data. See [Outlook verification](#outlook-verification).

Update `Documentation/ReleaseNotes.txt` for the selected version. Commit and
push submodule changes before committing their references in this repository.
After merging and pushing `main`, run each applicable release command:

```shell
DevelopmentTools\build.cmd release
DevelopmentTools\build.cmd publish
```

Test the release candidate before invoking either command. Extract the ZIP
to a separate directory and smoke-test it. Inspect the NuGet package's
dependencies and target frameworks and test it from a consumer project.
The library uses public MsgKit 3.0.6. Verify the packaged conversion
adapter preserves read-receipt fields using the public package. The fork is
no longer a submodule. See [MsgKit compatibility](#msgkit-mime-compatibility).

## Outlook verification

The [connection contract](OutlookLifecycle.md) describes behavior that
library consumers can rely on. The following coverage and procedures are
for maintaining that implementation.

### Automated coverage

Unit tests use the real coordinator with controlled operations and gates,
without manufacturing Outlook COM objects. They cover STA execution,
timeout validation, a blocked operation with a positive timeout, concurrent
factory callers sharing an outstanding probe, retries during cleanup, late
failure, and fresh probes after completion. Service tests cover reuse,
disposal, failure cleanup, reconnection, and compatibility with connections
that are not disposable.

Tests that activate Outlook or verify that an independent COM reference
survives disconnection are marked `Integration`. They require Outlook and
a configured interactive user session.

### Probe implementation history

The branch history explains why there are two application acquisitions:

- `587809c` acquired an application on a temporary STA and returned it.
- `8922de8` changed this into a boolean availability probe. The probe releases
  its application reference on its own thread; the real application is
  acquired on the caller's thread.
- `814f837` kept the probe reference local to the worker and used
  `Thread.Join` to avoid a late worker signaling an already-disposed event.

Do not optimize this into caching or returning the probe's COM wrapper.
The temporary STA exits, so that wrapper would depend on an apartment whose
lifetime has ended. The caller-thread acquisition preserves the reason for
the earlier RCW fix. It can attach to Outlook that the probe activated.

The probe releases its own application acquisition once, in its worker's
`finally` block. It does not call `Quit` or force every reference on a
potentially shared RCW to zero. No Outlook COM object leaves the probe.

### Manual regression checks

`ProbeApplication` lets errors propagate to `ProbeAttempt.Run`, which
captures and logs them even after the caller times out. A local catch that
only logs and returns would incorrectly report a successful probe.

Before release, exercise these scenarios with debug logging enabled:

| Initial state | Checks |
| --- | --- |
| Healthy Outlook running | Connect succeeds; repeated Connect reuses the session; Disconnect leaves Outlook usable. |
| Outlook closed with a configured profile | Connect activates Outlook and obtains a session; another Connect reuses it. |
| Original unresponsive Outlook state | Probe wait returns false; service has no connection or session; retries do not start additional blocked probes. |
| Recovery after a timeout | Restore Outlook and let the outstanding worker finish; retry in the same host succeeds. |

Use the same service and caller STA for each scenario's attempts. The
ordinary integration fixture starts Outlook during setup, so it cannot
establish the initially closed or unresponsive scenarios. Those require a
separate manual caller with no preliminary Outlook activation.

Record the commit, Windows/Outlook versions and bitness, profile setup,
exact reproduction steps, requested timeout, elapsed time, result,
`IsConnected`, session presence, and last logged stage. Preserve logs.
A repeat call after success should have no acquisition stages. Verify
Outlook remains usable after disconnect.

Use a disposable environment for deliberate failures. Give the manual test
host an external watchdog so an unbounded caller-thread COM call cannot
stall the exercise indefinitely. Watchdog expiry is an aborted test, not a
successful production timeout; terminating the host does not undo Outlook
side effects. Do not suspend or terminate an everyday Outlook session.

The original hung-Outlook scenario requires its actual reproduction steps.
A suspended process is a separate test case, not proof of reproducing that
failure. Unit tests validate coordination; healthy integration tests do not
validate the original failure. Record these scenarios separately rather
than treating a green unit suite as complete real-Outlook coverage.

### Optional VM testing

Start with one reusable Windows VM, a dedicated Outlook profile and test
data, and snapshots for repeatable baselines. Run it on demand before
releases; add Outlook-version or bitness variants only when useful.
A snapshot may still need explicit steps to reproduce a particular hang.

Keep the existing NUnit framework. A future CI job can restore the VM,
prepare a scenario, run explicitly selected integration tests, collect TRX
results and stage logs, then reset or shut down the VM. Keep an outer job
timeout and collect partial logs on failure. Implementing the scenario
runner and VM provisioning is deferred.

Plan for a configured, signed-in interactive desktop and an interactive
agent, rather than a Windows-service test runner. Microsoft's
[desktop test guidance](https://learn.microsoft.com/en-us/azure/devops/pipelines/test/ui-testing-considerations?view=azure-devops)
describes this setup. Its
[Office automation guidance](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa)
also explains why a VM does not eliminate Office automation limitations.

One VM used on demand limits compute costs. Licensing, snapshot storage,
Office updates, and profile maintenance still need budgeting. Keep fast,
deterministic unit tests in normal CI; the VM tier supplements them.

## Dependency regression checks

The library explicitly declares OpenMcdf `3.3.0` and, outside the .NET
Framework target, System.Security.Cryptography.Pkcs `10.0.13`. These security
minimums must appear in the generated package metadata so consumers do not
restore vulnerable versions through upstream MsgKit.

`MsgKitConversionTests` checks the production EML-to-MSG adapter without
Outlook. It covers optional IDs, inline references, attachment data, read
receipts, and source preservation. `MsgKitAttachmentTests` retains a valid
inline-attachment control against upstream MsgKit. Read-receipt data is
preserved as ordinary attachment content, without inventing Content-IDs.

## MsgKit MIME Compatibility

### Valid MIME and optional identifiers

RFC 2045 section 7 makes Content-ID generally optional. RFC 2183 section 2.1
uses `inline` to request automatic display; it does not require an HTML
`cid:` reference. RFC 2298 includes a read-receipt example whose notification
part has no Content-ID. Such parts are not inherently malformed.

An unmatched HTML `cid:` reference is a separate issue. RFC 2110 also allows
related parts to be identified by Content-Location. The original archived
message that triggered the constructor exception is still unavailable.

### Initial upstream comparison

Twelve synthetic EML cases were converted to MSG without Outlook. Checks
inspected attachment bytes, names, IDs, body content, and HTML references.
The inputs used period-compatible MIME structures and a 1998 date.

The local fork and upstream MsgKit 3.0.4 produced identical results, including
with upstream's MimeKit 4.16.0 dependency and the fork's MimeKit 4.18.2.
The comparison was subsequently repeated against upstream MsgKit 3.0.6.
Each configuration passed ten checks and failed two notification checks.

| Case | Unmodified converter result |
| --- | --- |
| Missing, blank, or unusable Content-ID | Attachment bytes retained |
| Missing filename | Bytes retained; fallback name supplied |
| Matching HTML cid | Reference, ID, and bytes retained |
| Unmatched HTML cid | Reference retained; no target ID invented |
| Relative Content-Location | Bytes/name/HTML retained; location tag absent |
| Read receipt with or without inline disposition | Notification data empty |

The notification tests verified that recipient and disposition fields were
present before conversion. The specialized MsgKit notification branch set a
filename but never wrote those fields into its attachment stream. This is a
data-preservation defect, separate from constructor validation.

The earlier constructor probe demonstrated that upstream rejects a direct
`Attachments.Add` call with `isInline: true` and a missing ID. The EML
converter avoids that condition for the examples above. Production
regressions therefore exercise conversion, rather than require the upstream
constructor to accept values it intentionally rejects.

### Email.ToolKit workaround

The library now uses public MsgKit 3.0.6 and MimeKit 4.18.2. `EmlConverter`
inspects the MIME body and replaces disposition-notification parts only in
an internal conversion copy. Their serialized fields become ordinary
attachments, defaulting to `read-receipt.txt`. Existing names and Content-IDs
are retained. No identifiers are generated.

The archived EML is never rewritten. Messages without notifications are
passed to MsgKit from the original input bytes. Attached messages remain
attached EML files; their internal MIME bodies are not normalized.

This preserves notification information as readable attachment data. It does
not create a native Outlook read-receipt object or restore unknown HTML
relationships. MIME inspection adds a parse to conversion; the inspected
message is disposed before MsgKit parses the selected input.

### Permanent regression coverage

`SourceCode/ToolKit.Tests/MsgKitConversionTests.cs` contains the original
conversion cases plus nested and standalone notifications, source-file
preservation, multiple receipts, existing metadata, and attached messages.
They run through the production adapter without Outlook.

The permanent baseline had five read-receipt failures. With the workaround,
all eighteen conversion cases pass against the public package. The valid
inline-constructor control also passes. No production build references the
fork project; its checkout is retained pending submodule cleanup.

Real Outlook import and rendering still require a release smoke test. These
synthetic cases do not establish compatibility with every archived message.
An upstream change is not required for the local workaround to function.

### Optional upstream report

The [unsent issue draft](#unsent-upstream-issue-draft) below contains
reproduction steps. `Examples/ReadReceiptWithoutContentId.eml` is synthetic
input.
The report concerns empty output data, rather than the earlier argument over
whether inline attachments need Content-IDs. No maintainer was contacted.

Initial comparison source and TRX results remain in the ignored local
`artifacts/msgkit-examination/` directory and examination ZIP.

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

Converting a multipart/report containing a message/disposition-notification
part creates an MSG attachment with an empty data stream. The input includes
Final-Recipient, Original-Message-ID, and Disposition fields. The human-readable
message body converts successfully, but the machine-readable fields are lost.

This reproduces with public MsgKit 3.0.6. It also reproduced with 3.0.4 and
with the local attachment-tolerance fork. No Outlook installation or COM
activation is needed to reproduce the conversion defect.

#### Input and reproduction

Use [the synthetic EML](Examples/ReadReceiptWithoutContentId.eml). Its missing
Content-ID is intentional: RFC 2298 section 9 includes the same general MIME
structure without a Content-ID on the notification part.

Create a console project targeting .NET 10 and reference MsgKit 3.0.6,
MimeKit 4.18.2, OpenMcdf 3.3.0, and System.Security.Cryptography.Pkcs 10.0.13.
Copy the EML into the working directory and run:

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

Start with no existing `receipt.msg`; the converter creates a new file.
Expected: a nonempty attachment containing the notification fields.
Actual: the printed attachment length is zero.

#### Apparent cause

In `Converter.ConvertEmlToMsg`, this branch sets the filename and exits:

```csharp
case MessageDispositionNotification notification:
    fileName = notification.FileName;
    break;
```

Unlike the delivery-status branch, it does not write the notification body
to `attachmentStream`. One possible upstream correction is to serialize its
content before exiting, then verify the decoded MSG attachment fields.

The regression should assert retained recipient, original-message ID, and
disposition values, both with and without an inline disposition header.
Checking only that conversion does not throw would miss the data loss.

#### Scope

This report concerns preservation of parsed notification data. It does not
request a change to attachment-constructor validation, or assume every
notification is referenced by an HTML cid URL.

Email.ToolKit has a local workaround that represents these fields as an
ordinary attachment in a conversion copy. An upstream correction would let
that workaround be retired once a fixed package is available and tested.

- [RFC 2298 example](https://www.rfc-editor.org/rfc/rfc2298.html#section-9)
- [MsgKit package](https://www.nuget.org/packages/MsgKit/3.0.6)
- [Current converter source](https://github.com/Sicos1977/MsgKit/blob/master/MsgKit/Converter.cs)

## Explicit publication

Configure a NuGet.org trusted-publishing policy with these values:

- Repository owner: `jamesjohnmcguire`
- Repository: `DigitalZenWorks.Email.ToolKit`
- Workflow filename: `build.yml`
- Environment: leave empty (the workflow does not declare one)

Commit and push the workflow and release changes to `main`. Complete the
release checks first, using Visual Studio MSBuild and the working Outlook
installation. The two distribution commands are independent:

```shell
DevelopmentTools\build.cmd release
DevelopmentTools\build.cmd publish
```

`release` builds the application, creates the executable ZIP with `zip.exe`,
and uploads only that ZIP to a draft GitHub release. Review and publish the
GitHub draft separately when ready. It does not pack the library or request
NuGet publication.

`publish` builds and packs the library locally. It uploads the NuGet package
and `package.commit` to the matching GitHub release, creating a draft if
needed. No executable build or ZIP is required. The GitHub release serves
as the package transfer location for the publishing workflow.

An existing draft must target the same commit. Each command replaces only
its own assets on retry, preserving the other distribution's artifacts.
For a published GitHub release, `publish` may add missing package assets
without replacing existing assets; the version tag must match the source
commit. Existing package assets are reused and verified by the workflow.
`release` refuses to replace an executable on a published GitHub release.

After uploading the package, the script dispatches `build.yml` from `main`
with the local package's SHA-256 and follows that specific run. The Ubuntu
job downloads the package and verifies the source commit, release target
or tag, package ID, version, and SHA-256 before obtaining a short-lived key
with `NuGet/login@v1` and pushing the exact package. No Office installation
or COM build occurs on the publishing runner. API-key arguments are no
longer supported.

The workflow runs only on explicit dispatch with `publish=true` from
`main`. Commit pushes, tags, and GitHub release events do not trigger it.
Its `contents: write` permission permits draft-release discovery; its steps
do not create, edit, or publish GitHub releases. Publication requests are
serialized. The workflow does not rerun tests or prove build provenance.

Both commands require `main` and a clean working tree including submodules.
They build their own artifact and verify its source commit before upload.
A build or upload failure stops the command before NuGet dispatch. If the
source commit changes, rebuild from the intended release commit and review
any existing draft target and version tag. Do not retarget a published
version.

Commands return a nonzero exit code on failure. If dispatch succeeds but
status tracking fails, inspect that run before retrying. Publication failure
details are available with `gh run view <run-id> --log-failed`.
