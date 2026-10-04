# AGENTS.md

## Repository Overview

Email.ToolKit is a .NET library for working with email data and Microsoft
Outlook.

The library is distributed as a NuGet package. Backward compatibility of
public APIs is therefore important.

Microsoft Outlook integration uses COM through
`Microsoft.Office.Interop.Outlook`. Outlook should be treated as an external,
optional service rather than as an always-available application dependency.

## General Development Guidelines

- Follow the existing project structure and conventions.
- Prefer existing abstractions before introducing new ones.
- Follow existing nullable-reference-type conventions.
- Avoid unnecessary dependencies.
- Keep changes narrowly scoped to the task being performed.
- Do not combine unrelated architectural refactoring with a focused bug fix
  or feature unless required.
- Investigate compiler and analyzer warnings rather than suppressing them
  without justification.
- Preserve existing public APIs unless a breaking change is explicitly
  intended.
- Remember that changes to public library APIs can affect existing NuGet
  consumers.

## C# Style

- Use Allman braces.
- Use tabs for indentation.
- Prefer one statement per line.
- Use UTF-8 and LF line endings.
- Follow `.editorconfig` where applicable.
- Prefer explicit, readable code over unnecessarily compact code.
- Follow the existing error-handling conventions.
- Prefer existing result/error abstractions where applicable rather than
  introducing competing patterns.

## Outlook Architecture

### OutlookService

`OutlookService` is the primary public entry point for new Outlook-related
functionality.

Clients should not normally need to know how Outlook is discovered, started,
or connected.

`IOutlookService` should represent the public service contract.

The normal public connection API should remain simple, for example:

    outlookService.Connect();

Implementation dependencies such as `IOutlookFactory` should not be required
by normal callers.

### OutlookFactory

`OutlookFactory` is responsible for obtaining Outlook and creating the
infrastructure required for a connection.

Responsibilities may include:

- Determining whether Outlook can be activated.
- Attempting to obtain an already-running Outlook instance.
- Creating an Outlook application when necessary.
- Creating an `IOutlookConnection`.

Factory abstractions primarily exist to separate COM creation from service
logic and to make service behavior testable.

Do not expose factory infrastructure through the normal public
`IOutlookService` API without a specific reason.

### OutlookConnection

`OutlookConnection` wraps the underlying Outlook COM application.

Keep `OutlookConnection` internal unless there is a deliberate architectural
reason to expose it.

The raw `Microsoft.Office.Interop.Outlook.Application` object should remain
an implementation detail of the Outlook connection infrastructure.

A connection should provide a stable session wrapper rather than creating a
new `OutlookSession` on every access.

Connection lifecycle responsibilities, including Outlook ownership and COM
cleanup, should increasingly be encapsulated by the connection layer rather
than leaked into higher-level callers.

### OutlookSession

Prefer `IOutlookSession` at abstraction boundaries.

Avoid exposing raw Outlook COM objects through new public APIs when a library
abstraction can reasonably represent the required functionality.

Some existing APIs expose Outlook COM types for historical compatibility.
Do not expand this exposure unnecessarily.

## COM Boundaries

Treat `Microsoft.Office.Interop.Outlook` types as implementation details where
practical.

Code using types such as:

- `Outlook.Application`
- `Outlook.NameSpace`
- `Outlook.Store`
- `Outlook.MAPIFolder`
- `Outlook.MailItem`

should generally remain close to the Outlook COM integration layer.

Prefer the namespace alias:

    using Outlook = Microsoft.Office.Interop.Outlook;

and explicit names such as:

    Outlook.Application

in COM-facing code. This makes COM dependencies visually apparent.

Do not introduce direct Outlook COM dependencies into higher-level classes
without a clear reason.

## Outlook Lifetime and Threading

Outlook COM lifetime and threading require special care.

Do not assume that an Outlook process was started by this library merely
because an Outlook COM connection was successfully obtained.

Distinguish between:

- attaching to an Outlook instance that was already running; and
- starting or activating Outlook on behalf of the library.

Do not terminate an independently started Outlook instance during normal
cleanup.

Do not use process detection alone as proof of COM ownership.

Outlook COM operations may require STA-thread handling. Avoid moving COM
objects across thread boundaries without understanding the RCW and apartment
lifetime implications.

Timeout code that stops waiting for an STA operation does not necessarily
cancel the underlying COM operation.

Long term, Outlook should be treated as an external service with explicit
connection, ownership, threading, and lifetime management.

## Legacy OutlookAccount

`OutlookAccount` is legacy singleton-based infrastructure.

Preserve it where necessary for backward compatibility, but avoid introducing
new dependencies on `OutlookAccount.Instance`.

New Outlook functionality should generally move toward:

    OutlookService
        -> IOutlookConnection
        -> IOutlookSession

Do not introduce adapters back to `OutlookAccount.Instance` merely to make
new abstractions compile unless that compatibility bridge is deliberately
required.

## Migrate

`Migrate` contains legacy static migration APIs that are part of the existing
NuGet surface.

Preserve those APIs during 1.x maintenance unless a breaking change is
explicitly intended.

Do not force the current Outlook service/session architecture through
`Migrate` as part of unrelated work if doing so requires widespread API or
behavioral changes.

`Migrate` is a candidate for substantial redesign in a future major version.
Future migration functionality may support destinations other than Outlook
PST files, so avoid unnecessarily coupling new migration concepts specifically
to PST or Outlook.

## Testing

Use NUnit for tests.

Prefer deterministic unit tests for service and orchestration behavior.

Use fakes at established abstraction boundaries such as `IOutlookFactory`,
`IOutlookConnection`, and `IOutlookSession`.

Do not create fake Outlook COM objects merely to satisfy unit tests.

Do not change production methods to `virtual`, widen visibility, or introduce
inheritance solely to manufacture a test seam unless there is a strong design
reason.

A test should exercise meaningful production behavior. Avoid tests that only
verify behavior hard-coded into the fake itself.

Call-count properties on fake objects are test instrumentation. Use them when
the number of dependency calls is meaningful, such as verifying that an
expensive Outlook connection is not repeatedly created.

Distinguish unit tests from Outlook integration tests.

Tests involving real PST creation, Outlook COM behavior, Outlook profiles, or
actual `Outlook.Application` activation are integration tests even if they
currently reside in the same NUnit project.

Do not replace integration tests with fakes when the behavior being tested
inherently depends on real Outlook/PST/COM functionality.

## Refactoring Guidance

Prefer incremental refactoring with a green test suite between steps.

When encountering legacy architecture during a focused change:

1. Identify the architectural problem.
2. Record or document it when useful.
3. Fix it only if necessary for the current task.
4. Avoid allowing one refactoring to cascade through unrelated working code.

Do not blindly propagate a new abstraction through the repository simply
because the abstraction exists.

If new code requires casts from an interface back to a concrete
implementation, consider whether the abstraction is incomplete before
spreading the cast through additional code.

Prefer stable intermediate states over large partially completed architectural
migrations.

## Public API Compatibility

Treat the existing NuGet API surface as a compatibility constraint.

Before removing, renaming, changing visibility, changing parameters, or
changing return types of an existing public member, determine whether it is
part of the released package API.

When obsolete public functionality must remain for compatibility, prefer a
deprecation/obsolete path over immediate removal.

Breaking API redesigns should normally be reserved for an intentional major
version change.