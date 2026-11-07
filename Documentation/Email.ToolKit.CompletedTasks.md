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
