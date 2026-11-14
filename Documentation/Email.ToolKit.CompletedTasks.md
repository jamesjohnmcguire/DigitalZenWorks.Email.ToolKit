# Email.ToolKit Completed Tasks

This records completed work and superseded proposals moved out of the
[active backlog](Email.ToolKit.Tasks.md). It is not an additional task list.
Partial historical findings do not establish coverage for every input.

## Prepared for 1.14.1 (not yet published)

- [x] Add failing tests for the exact released signatures below, including
  return types; restore forwarding overloads with historical behavior:
  - `Migrate.EmlFileToPst(string, string)`.
  - `Migrate.EmlToPst(string, string, bool)`.
  - `OutlookFolder.AddFolder(MAPIFolder, string)`.

- [x] Move shared test helpers and embedded fixtures into `Support`, keeping
  existing resource names and bytes.

- [x] Retain two-argument `AddFolder` as a supported convenience overload,
  without an obsolete annotation; forward with normalization disabled.

- [x] Review all 20 existing obsolete annotations and correct replacement
  guidance. Add warning-only annotations to the two restored migration
  overloads, preserving the supported two-argument AddFolder overload.
  Document consumer warning implications and version 2 considerations.
  Validation: Debug build and 182 selected tests pass (178 deterministic,
  four real-Outlook cases); production method implementations are unchanged.

- [x] Rebuild the candidate with the reviewed obsolete annotations and
  pass ApiCompat against published 1.12.62 and 1.14.0.
- [x] Validate consumer compilation and unchanged-binary binding using the
  PackageClients EmailToolKitClient on net472, net10.0 (netstandard2.0
  asset), and net10.0-windows. Nine builds and eighteen successful runs
  passed. Three negative controls failed as expected: unchanged 1.12.62
  clients cannot bind the missing methods in published 1.14.0.
  The candidate restores binding for those clients and preserves binding
  for unchanged 1.14.0 clients. Recompiled legacy calls emit exactly two
  intended obsolete warnings per target. Existing unit/integration
  validation remains 178 deterministic and four focused Outlook cases.
- [x] Add reusable Current/Legacy API modes and instructions to the private
  PackageClients repository. No package publication or version bump was
  performed during validation.

- [x] Finalize 1.14.1 release notes and validate the candidate on 2026-11-14: 488 full-solution tests, 289 Release ToolKit tests, and 38 packaged MIME regression cases passed. Three explicit/manual Outlook cases were not run. The local package built for all eight target frameworks with embedded PDBs and XML documentation; ApiCompat passed against 1.12.62 and 1.14.0. Twelve consumer target builds and eighteen runtime checks passed, including six unchanged-binary checks. The extracted executable ZIP passed help and invalid-timeout smoke checks. Validation artifacts used command-line 1.14.1 version overrides; project version tags remain unchanged. No distribution was published.

- [x] Complete the final dependency refresh and release-day test hardening: retain MimeKit 4.18.2 for MsgKit 3.0.6 compatibility, give real test PSTs GUID-unique filenames, and preserve backing files if detachment fails. The refreshed candidate passed all 488 solution tests from a fresh non-admin Outlook session, 38 packaged MIME cases, both API baselines, twelve consumer target builds, eighteen runtime checks, executable ZIP smoke checks, and the eight-framework vulnerability audit. Three explicit/manual tests were not run. No sleeps, automatic Outlook restarts, library API changes, project version bump, or publication were introduced by this test-support work. Repeated-run resource retention remains follow-up work.

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
