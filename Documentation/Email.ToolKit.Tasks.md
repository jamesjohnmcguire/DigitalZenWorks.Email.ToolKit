# Email.ToolKit Tasks

This is the active backlog after 1.14.0. Checkboxes identify remaining work,
not release commitments. Historical reports are not assumed fixed merely
because their implementation changed.

See [maintainer notes](DevelopmentNotes.md) for release procedures,
[Outlook lifecycle](OutlookLifecycle.md) for the design, and the
[API audit](ApiCompatibilityAudit.md) for compatibility findings.

Completed work and superseded proposals are in
[Email.ToolKit.CompletedTasks.md](Email.ToolKit.CompletedTasks.md). Keep this
file focused on unfinished work; move completed entries to that record.

## Next maintenance patch: 1.14.1

No affected consumers have been identified. Prepare this patch at a normal
maintenance pace. Historical release intervals do not establish whether
those releases addressed bugs or compatibility.

- [ ] Finalize release notes. Near release, update application/library versions
  together in one commit; retain `AssemblyVersion` for compatible changes.
  Follow the maintainer checklist for the separate executable and NuGet
  distributions.

## Testing and maintainability

- [ ] Add focused `NormalizeFolderName` tests, including its regex cases.
- [ ] Review test naming and remaining `Test` prefixes. Group fixtures by
  subject when useful; keep methods alphabetical and execution independent.
- [ ] Reuse cleanup helpers where useful without hiding test behavior.
- [ ] Identify real-Outlook fixtures consistently. Keep ordinary integration
  tests in the normal test script; choose `[Explicit]` for manual scenarios
  individually rather than disabling every integration test by default.
- [ ] Review intended package architecture and transitive dependencies.
  Published 1.12.62/1.14.0 library DLLs target x64; the local candidate is
  managed AnyCPU. Consumer checks currently cover 64-bit Windows. Verify
  actual 32-bit loading before promising support for it.
- [ ] Test without Outlook installed. Distinguish unavailable-service runtime
  behavior from the separate COM-reference build requirement.
- [ ] Enable nullable checking project-wide incrementally, with tests for
  intended input contracts; preserve legacy public contracts in 1.x.
- [ ] Review exception conventions and when methods should throw versus
  return a failure result. Preserve established 1.x behavior unless a change
  is explicitly intended and covered by tests.
- [ ] Review spacing after file-scoped namespace declarations (the requested
  wider separation), keeping formatting changes separate from behavior fixes.
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
- [ ] Deferred: add a client-validation callback to CommandLine.Commands.
  Once available, consider moving application `--timeout` validation into
  that callback. Retain current validation until its replacement is tested.
- [ ] Review old CommandLineCommands bug notes against specific reproductions;
  its publication setup is already established.
- [ ] Optionally submit the existing MsgKit read-receipt data-loss issue
  draft in [maintainer notes][msgkit-draft]. It remains unsent; the local
  workaround does not depend on upstream action.
- [ ] Reevaluate the adapter after a suitable upstream fix and passing tests.
  Former fork-maintenance PRs are not prerequisites for releases.

[comparison-reference]: Email.ToolKit.DuplicateComparison.md

[msgkit-draft]: DevelopmentNotes.md#unsent-upstream-issue-draft

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

New
	Organize Test Files	I recommend subject folders—CommandLine, Conversion, Outlook, and Compatibility—and incrementally splitting the large EmailToolKitTests fixture by behavior. Keep methods alphabetical and independent of execution order.
