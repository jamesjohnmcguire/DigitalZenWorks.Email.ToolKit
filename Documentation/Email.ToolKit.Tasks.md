# Email.ToolKit Tasks

This backlog contains unfinished work, not release commitments. Remove completed entries after recording significant changes in [release notes](ReleaseNotes.txt) and durable decisions in the existing reference documents. Routine validation logs do not need a permanent task record.

See [maintainer notes](DevelopmentNotes.md), [Outlook lifecycle](OutlookLifecycle.md), the [API audit](ApiCompatibilityAudit.md), and [duplicate comparison](Email.ToolKit.DuplicateComparison.md).

## Release tooling

- [ ] Keep automatic tagging consistent between `build.cmd release` and `build.cmd publish`. Verify an existing tag identifies the intended commit; never move a published tag. Ensure release notes are finalized before upload.
- [ ] Decide whether `build.cmd` needs a local-only Release build option; the current `release` command also uploads the executable.
- [ ] Evaluate removing `IncludeAllContentForSelfExtract` to resolve NETSDK1244. Validate the published executable's Outlook operations before changing extraction behavior.

## Testing and maintainability

- [ ] Add focused `NormalizeFolderName` tests, including regex cases.
- [ ] Organize fixtures by subject (`CommandLine`, `Conversion`, `Outlook`, `Compatibility`), splitting `EmailToolKitTests` incrementally without changing setup/cleanup behavior. Review remaining `Test` prefixes; keep methods alphabetical and execution independent.
- [ ] Reuse cleanup helpers where useful without hiding test behavior. Identify real-Outlook fixtures consistently; keep ordinary integration tests in `test.cmd`, with `[Explicit]` reserved for individual manual scenarios.
- [ ] Review package architecture and transitive dependencies. Published 1.12.62/1.14.0 library DLLs target x64; the 1.14.1 candidate was managed AnyCPU, with consumer checks on 64-bit Windows. Verify actual 32-bit loading before promising support.
- [ ] Test runtime behavior without Outlook installed, separately from the COM-reference build requirement.
- [ ] Enable project-wide nullable checking incrementally, testing intended input contracts and preserving legacy public contracts in 1.x.
- [ ] Review exception-versus-failure-result conventions. Preserve established 1.x behavior unless changes are intentional and tested.
- [ ] Review the requested wider spacing after file-scoped namespaces separately from behavior changes.
- [ ] Review analyzer warnings and boolean-return documentation. Remove interim security references only when upstream dependencies retain the required fixes.
- [ ] Review PST-path handling outside `OutlookSession` before introducing another normalization helper.
- [ ] Clarify delayed merge/move processing, folder paths, merge rules, command defaults, and version output for invalid-input versus explicit help.
- [ ] Extend appointment/contact hash inventories and MAPI mappings in the duplicate comparison reference. Test body normalization, read-state differences, and repeated-message send times before changing comparison rules.

## Investigations needing reproduction

Record the version, input, Outlook state, reproduction steps, and expected result before treating these reports as confirmed defects.

- [ ] Investigate cumulative Outlook/PST resource retention during repeated integration runs. Review COM ownership and reduce PST churn where isolation permits; retain independent creation/attachment tests. See [operational findings](DevelopmentNotes.md#outlookpst-resource-retention). Persistent locks are not evidence that a short pause will fix the issue.
- [ ] Recheck attaching to running Outlook across elevation levels, profiles, and installation/bitness combinations.
- [ ] Reproduce deletion failures involving particular item classes, special folders, and read-only calendars.
- [ ] Investigate attachment filename/save failures and temporary-file cleanup.
- [ ] Obtain sanitized historical EML samples to supplement synthetic attachment tests; verify real Outlook import and rendering.
- [ ] Recheck slash-containing folder names, relative PST paths, repeated/digit-suffixed names, and merges that appear to do nothing.
- [ ] Investigate mangled-message and missing-recipient reports using original DBX data.
- [ ] Improve merge/progress output with item/folder counts, skipped types, and cases with nothing to process.
- [ ] Reproduce remaining missing-method/test reports; distinguish stale assemblies from API changes using the historical audit.
- [ ] Review logging when a consuming application supplies no configuration.

## Version 2 design candidates

These require design review; they do not authorize breaking changes in 1.x.

- [ ] Design a long-lived Outlook STA owner with COM dispatch, shutdown, reentrancy, timeout, and late-work rules. Stopping a wait does not cancel COM work.
- [ ] Reduce public COM exposure; consider an Outlook-specific library and review class names after deciding package boundaries.
- [ ] Redesign migration around explicit dependencies and possible non-PST destinations without forcing that redesign through the legacy static `Migrate` contract.
- [ ] Separate EML migration/copy helpers and folder/store/session responsibilities. Review `AddMsgFile` layering, directory enumeration, and proposed public helpers individually.
- [ ] Consider a success result for `DbxDirectoryToPst` as an intentional major-version change; retain released void signatures in 1.x.
- [ ] Review obsolete wrappers, account-based constructors, singleton use, store-removal methods, and ID naming. Require usable migration paths before further deprecation/removal.
- [ ] Define profile selection and ownership/lifetime of returned Outlook objects.

## Feature candidates and related projects

- [ ] Continue profile support exploration from branch `profile-info`; determine its release scope separately from the version 2 ownership design.
- [ ] Broaden contacts/calendar/report migration and duplicate detection.
- [ ] Evaluate MBOX, Novell, MsMail, and EMC import requirements; some belong to other email projects.
- [ ] Consider configurable reserved folders, sender-based moves, attachment export, folder sorting, and Outlook rules.
- [ ] Evaluate UI, installer, add-in, configuration, and promotion work against demand. Justify native MAPI/performance work with measurements.
- [ ] Review DBX file/header validation, tree ordering/indexes, iteration, interim folders, and malformed-message coverage in DbxOutlookExpress; track implementation there.
- [ ] Revisit a client-validation callback in CommandLine.Commands. Move `--timeout` validation only after a tested replacement exists; check old command-line bug notes against specific reproductions.
- [ ] Optionally submit the [MsgKit read-receipt issue draft](DevelopmentNotes.md#unsent-upstream-issue-draft). Retire the local adapter workaround only after a suitable upstream fix passes regression tests.

## Historical research

The adjacent Email repository retains the original `Email/Support/Email.ToolKit.Tasks.txt` snapshot and broader email/native-MAPI research. Unreviewed entries are not assumed resolved. The maintained duplicate comparison reference contains the curated property inventory and mappings; historical snippets do not specify current hash behavior.
