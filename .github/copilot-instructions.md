# Repository Instructions

## Coding style

- Use Allman brace style.
- Use tabs for indentation.
- Keep one statement per line.
- Prefer explicit, readable code over clever or overly compact code.
- Preserve UTF-8 encoding and LF line endings.
- Follow the existing `.editorconfig`.
- Do not suppress analyzer warnings unless there is a clear justification.

## .NET

- Treat compiler and analyzer warnings as issues to investigate.
- Prefer existing project abstractions before introducing new ones.
- Follow existing nullable-reference-type conventions.
- Avoid unnecessary dependencies.

## Error handling

- Prefer the repository's existing `Result<T>` pattern where applicable.
- Do not replace structured error handling with broad exception catching.
- Avoid `catch (Exception)` unless specifically justified.

## Testing

- Add or update tests when changing behavior.
- Follow existing NUnit conventions for C# tests.
- Do not weaken tests merely to make them pass.
- Run the relevant tests after making changes.

## Changes

- Keep changes narrowly scoped to the requested task.
- Avoid unrelated cleanup unless it is required for the change.
- Before making architectural changes, explain the proposed approach.
- Preserve existing public APIs unless changing them is part of the task.
