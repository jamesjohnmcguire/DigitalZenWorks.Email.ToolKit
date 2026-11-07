# Email.ToolKit Tasks

This is the active backlog after 1.14.0. Checkboxes identify remaining work,
not release commitments. Historical reports are not assumed fixed merely
because their implementation changed.

See [maintainer notes](DevelopmentNotes.md) for release procedures,
[Outlook lifecycle](OutlookLifecycle.md) for the design, and the
[API audit](ApiCompatibilityAudit.md) for compatibility findings.

## Next maintenance patch: 1.14.1

No affected consumers have been identified. Prepare this patch at a normal
maintenance pace. Historical release intervals do not establish whether
those releases addressed bugs or compatibility.

- [ ] Add failing tests for the exact released signatures below, including
  return types; restore forwarding overloads with historical behavior:
  - `Migrate.EmlFileToPst(string, string)`.
  - `Migrate.EmlToPst(string, string, bool)`.
  - `OutlookFolder.AddFolder(MAPIFolder, string)`.
- [ ] Retain two-argument `AddFolder` as a supported convenience overload,
  without an obsolete annotation; forward with normalization disabled.
- [ ] Audit obsolete messages for accurate names, signatures, replacements,
  and migration guidance. Consider warning-only annotations for restored
  migration overloads after confirming their replacements. Account for
  consumers treating warnings as errors.
- [ ] Compare the candidate package against 1.12.62 and 1.14.0 with ApiCompat.
  Check consumer compilation and binary binding; run relevant unit and
  real-Outlook integration tests.
- [ ] Update release notes. Near release, update application/library versions
  together in one commit; retain `AssemblyVersion` for compatible changes.
  Follow the maintainer checklist for the separate executable and NuGet
  distributions.

## Testing and maintainability

- [ ] Add focused `NormalizeFolderName` tests, including its regex cases.
- [ ] Review test naming, remaining `Test` prefixes, and support-file layout.
  Keep tests alphabetical when editing them.
- [ ] Reuse cleanup helpers where useful without hiding test behavior.
- [ ] Identify real-Outlook fixtures consistently. Keep ordinary integration
  tests in the normal test script; choose `[Explicit]` for manual scenarios
  individually rather than disabling every integration test by default.
- [ ] Test without Outlook installed. Distinguish unavailable-service runtime
  behavior from the separate COM-reference build requirement.
- [ ] Enable nullable checking project-wide incrementally, with tests for
  intended input contracts; preserve legacy public contracts in 1.x.
- [ ] Review analyzer warnings and boolean-return documentation. Reevaluate
  interim transitive security references when updating dependencies; remove
  them only after confirming the needed fixes remain.
- [ ] Review PST-path handling outside `OutlookSession`. Its existing helper
  and tests cover session path selection; identify inconsistent callers
  before introducing another normalization helper.
- [ ] Clarify delayed merge/move processing, folder paths, merge rules, and
  command defaults. Check version output for invalid-input help as well as
  explicit help.
- [ ] Maintain the [duplicate comparison reference][comparison-reference],
  extending appointment/contact inventories and MAPI mappings. Evaluate body
  normalization, read/unread differences, and repeated-message send times
  with regression cases before changing duplicate detection.

## Investigations needing reproduction

These are retained reports, not confirmed defects in 1.14.0. Record the
version, input, Outlook state, steps, and expected result before closing or
fixing them. Original details remain in the historical notes.

- [ ] Recheck attaching to running Outlook, including elevated versus normal
  execution, profiles, and installation/bitness details.
- [ ] Capture intermittent Outlook hangs, PST-load failures, and widespread
  test failures cleared by restarting Outlook. Use stage diagnostics and
  manual procedures; not every failure is a timeout bug.
- [ ] Reproduce deletion failures involving particular item classes,
  reserved/special folders, and read-only calendars; identify actual types.
- [ ] Investigate attachment filename/save failures and temporary-file cleanup.
- [ ] Obtain sanitized historical EML samples alongside synthetic attachment
  tests; verify real Outlook import and rendering.
- [ ] Recheck slash-containing folder names, relative PST paths, repeated or
  digit-suffixed names, and merges that appear to do nothing.
- [ ] Investigate mangled-message and missing-recipient reports using the
  original DBX data; unrelated MIME tests do not establish a fix.
- [ ] Improve confusing merge/progress output with useful item/folder counts,
  skipped item types, and cases with nothing to process.
- [ ] Revisit old missing-method/test failures with current reproductions;
  distinguish stale assemblies from actual API changes using the API audit.
- [ ] Review logger behavior when a consuming application supplies no setup.

## Version 2 design candidates

These require design review, not automatic implementation or removal in 1.x.

- [ ] Design a long-lived Outlook STA owner and dispatch all COM work onto it.
  Define shutdown, reentrancy, timeouts, and late-work behavior; stopping a
  wait must not imply cancellation of the COM operation.
- [ ] Reduce public COM exposure; consider an Outlook-specific library.
  Review class names after deciding package boundaries.
- [ ] Redesign migration around explicit dependencies and possible non-PST
  destinations. Do not force constructor injection into the existing static
  `Migrate` contract during a compatibility patch.
- [ ] Untangle EML migration/copy helpers and folder/store/session duties.
  Review `AddMsgFile` layering, directory enumeration, and any proposed
  newly public helpers individually.
- [ ] Consider a success result for `DbxDirectoryToPst` as an intentional
  major-version change; keep released void signatures throughout 1.x.
- [ ] Review obsolete wrappers, account-based constructors, singleton use,
  store-removal methods, and ID naming for replacement/removal. Require a
  usable migration path before deprecating additional APIs.
- [ ] Review Outlook profiles and ownership/lifetime of returned objects.

## Other feature candidates and related projects

Prioritize separately from the compatibility patch.

- [ ] Broaden contacts/calendar/report handling, including migration and
  duplicate detection.
- [ ] Evaluate MBOX, Novell, MsMail, and EMC import requirements separately;
  some belong to other email projects.
- [ ] Consider configurable reserved folders, sender-based message moves,
  attachment export, folder sorting, and Outlook rules.
- [ ] Evaluate UI, installer, add-in, configuration, and promotion work
  against demand. Justify native MAPI/performance work with measurements.
- [ ] Review DBX parser coverage in DbxOutlookExpress: file/header validation,
  tree ordering/indexes, iteration, interim folders, and malformed messages.
  Track implementation in that repository.
- [ ] Review old CommandLineCommands bug notes against specific reproductions;
  its publication setup is already established.
- [ ] Optionally submit the existing MsgKit read-receipt data-loss issue
  draft in [maintainer notes][msgkit-draft]. It remains unsent; the local
  workaround does not depend on upstream action.
- [ ] Reevaluate the adapter after a suitable upstream fix and passing tests.
  Former fork-maintenance PRs are not prerequisites for releases.

[comparison-reference]: Email.ToolKit.DuplicateComparison.md

[msgkit-draft]: DevelopmentNotes.md#unsent-upstream-issue-draft

## Completed or superseded entries

These dispositions do not establish coverage for every historical input.

| Former task | Disposition |
| --- | --- |
| Migrate to MsgKit NuGet; remove fork | Completed: public package referenced; fork submodule removed. |
| Avoid loose-attachment failures without fork | Synthetic conversion coverage exists; archived samples remain follow-up work. |
| Preserve read-receipt fields | Implemented in `EmlConverter`, with regression tests. |
| Add RemoveStore tests | Null, missing, unattached, detached, and extensionless paths are covered. |
| Add common integration cleanup | `OutlookTestContext` provides shared tracking and cleanup. |
| Is MapiItem deprecated? | All nine public methods already have obsolete annotations. |
| Replace List-based date/time APIs | Read-only overloads and obsolete wrappers exist; the historical IList gap is documented separately. |
| Add Outlook requirement to README | Present. |
| Add PST-path helper | Session helper/tests exist; review other callers separately. |
| Name imported PST stores | Migration code assigns root-folder names; broader store naming remains unverified. |
| Refactor Program.cs | Command wrappers moved to `Commands`; further work needs specific objectives. |
| Old MsgKit project package/XML/build warnings | Superseded by fork removal; current dependency warnings need their own checks. |
| Review scripts; publish NuGet and ZIP | Procedures documented and 1.14.0 published; verification remains a recurring release responsibility. |
| Deprecate two-argument AddFolder | Rejected: retain the useful convenience overload. |

## Historical source and scope

The original list remains an inactive snapshot in the adjacent Email
repository at `Email/Support/Email.ToolKit.Tasks.txt`. Unreviewed entries are
not assumed resolved merely because they are absent from this curated list.

`EmailDevelopmentNotes.txt` also covers EMC, MsMail, MAPI/C++, debugging, and
research. `MapiCompareMessagesSnippets.txt`, `MapiProperties.txt`,
`Properties.txt` remain research inputs for the maintained duplicate
comparison reference. `Dbx2Pst.txt` remains broader historical reference
material there; none of these files alone specifies current hash behavior.

Outlook conclusions are consolidated in
[the lifecycle design](OutlookLifecycle.md#future-sta-ownership-design).
Release procedures and MsgKit conclusions are in the maintainer notes.
The old planning and release files point to the maintained documents.
