# Outlook connection lifetime and probe timeout

This document is for developers using the NuGet library. It explains class
relationships, connection behavior, and the implementation constraints that
affect callers. Build, test, and release procedures are in the
[maintainer notes](DevelopmentNotes.md).

This is the interim design for `OutlookService`, `OutlookFactory`, and
`OutlookConnection`. It isolates Outlook COM access while preserving the
existing public interfaces. It does not replace the legacy `OutlookAccount`
singleton throughout the library.

## API relationships

Applications normally connect through `IOutlookService`, implemented by
`OutlookService`, then obtain its `IOutlookSession`. `OutlookSession` is the
default session wrapper. Create store and folder helpers with
`OutlookStore.FromSession(session)` and `OutlookFolder.FromSession(session)`.
Both factories accept `IOutlookSession` but currently require a concrete
`OutlookSession`, as supplied by `OutlookService`. Other implementations
are rejected with `NotSupportedException`.

The released account-based constructors remain public. New session and raw
COM constructors are internal. Some older methods still expose COM types
for compatibility. Connection infrastructure remains an implementation detail.

`Migrate` and `OutlookAccount.Instance` remain separate legacy entry points
for compatibility. Creating or configuring a service does not replace the
singleton used internally by those migration APIs.

## Why retain this interim design

The timeout request became more important when a running but unresponsive
Outlook left activation blocked. Process detection could not establish
whether Outlook was usable. The temporary STA probe bounds the caller's
wait for that activation without returning a wrapper from a dying apartment.

Failed probes leave the service disconnected. Retries share an outstanding
probe worker, successful connections are reused, and disconnect leaves
shared Outlook running. The second, caller-thread acquisition remains
unbounded while the larger apartment-ownership redesign is deferred.

## Reuse a service for one workflow

Create one `OutlookService` for a workflow and use it on the same STA thread
for connection, session operations, and disconnection. Once connected,
repeated `Connect()` calls reuse the connection and its stable session
wrapper. They do not probe Outlook or acquire another application reference.
A new service has an independent connection lifecycle.

The factory created by public `Connect()` is lightweight. It is not an
application cache. There is no process-wide service singleton and no promise
that one service can be used concurrently or moved between apartments.

```csharp
// Execute this workflow on an STA thread.
IOutlookService service = new OutlookService();

try
{
	if (service.Connect())
	{
		IOutlookSession? session = service.Session;
		// Perform this workflow's session operations here.
	}
}
finally
{
	service.Disconnect();
}
```

`IsConnected` means that the service has a connection and a session. It is
not an ongoing health check for an external Outlook process.

## Keep the temporary probe separate from the real connection

The probe activates Outlook on a temporary STA and releases that thread's
application reference during cleanup. No Outlook COM object is returned
from that worker, because its apartment exits when the probe completes.

The real application and session are acquired on the caller's STA. They can
attach to Outlook activated by the probe, but that acquisition is outside
the requested probe wait. Reuse that caller thread for the workflow.

The probe does not call `Quit` or forcibly release other clients' COM
references. This protects shared Outlook use while checking availability.

## Ownership and disconnect

Obtaining an application does not establish authority to shut down Outlook.
Neither finding a process nor falling back to `new Outlook.Application()`
proves exclusive ownership. The probe itself might have started Outlook,
and another client may also be using that application.

Factory-created connections therefore have shared or unknown ownership.
Normal `Disconnect()` no longer calls `Quit()`. This is an intentional
behavior change: Outlook may remain open after the workflow, including when
the workflow caused activation.

The internal connection implements `IDisposable`. Disposal clears its
managed application and session references without forcing COM release.
Existing APIs can expose session wrappers and raw COM objects to callers;
forcing an RCW's reference count to zero could break those other holders.
The runtime performs eventual COM release once the wrappers are no longer
referenced. This interim policy does not guarantee prompt COM cleanup or
Outlook process exit, and borrowed objects are not forcibly invalidated.

The service invokes disposal when the connection supports it, then clears
its own state in a `finally` block. Repeated disconnects do nothing further.
A disposal exception still propagates, but the service can reconnect.
Older implementations of `IOutlookConnection` need not implement a new
interface member: its public contract is unchanged.

The existing `IOutlookConnection.Quit()` remains available for compatibility.
It is an explicit application shutdown request, separate from disposal.
Only callers with authority to close the shared application should use it.

## What the timeout bounds

`timeOutSeconds` is the requested wait for the temporary probe thread to
finish activation **and cleanup**. Supported values are 0 through 2147483
seconds, the whole-second range supported by `Thread.Join(TimeSpan)`.
Validation occurs before a new probe is started. Zero requests no wait;
the result depends on whether the attempt has already completed.

| Stage | Covered by the requested probe wait? |
| --- | --- |
| Temporary STA activation and its cleanup | Yes |
| Joining a previously timed-out probe | Yes, with this caller's own wait |
| Scheduling, coordinator locking, and logging | No strict wall-clock bound |
| Caller-thread application acquisition | No |
| Session construction and later COM calls | No |

This is not an end-to-end deadline for `Connect()`. Scheduling and other
work can also make the method take longer than the requested wait.

Timing out stops waiting; it does not cancel the COM operation. The worker
may activate Outlook later and must still release its own acquisition.
The service returns false without publishing a connection. A timeout log is
distinct from an activation or cleanup failure; it does not prove Outlook
is unavailable.

## Retries and outstanding workers

All default factories in the same loaded library share one probe
coordinator. While its worker is alive, including cleanup, another caller
joins that attempt instead of creating another worker. This also applies
when public `Connect()` constructs a fresh factory for a retry.

Each caller retains the attempt it joined and has its own wait. A later
attempt cannot overwrite that result. Once a worker exits, a subsequent
caller starts a fresh probe; success and failure are not permanent caches
of Outlook availability. Coordination is not shared across processes or
separate loads of the library.

A permanently blocked COM call leaves one outstanding probe. Further calls
can time out on that same worker; they cannot cancel or replace it. This
bounds outstanding probe growth, not COM execution time. The coordinator
does not serialize caller-thread connection acquisition or session use.

Worker failures are captured and logged even when no caller is still
waiting. COM and invalid-operation failures produce false for waiting
callers. Unexpected failures are rethrown to callers that observe
completion, preserving the exception. Late failures must not become
unhandled background-thread exceptions that terminate the host process.

## Design limits and future work

A strict overall deadline requires a larger design: a long-lived STA owner,
dispatch of all COM work to that apartment, and explicit rules for timed-out
work and cleanup. A deadline still would not make a hung COM call safely
cancellable. That redesign, deterministic lifetime management of all
borrowed wrappers, and any verified exclusive-shutdown policy are deferred.

## Connection diagnostics

Debug logs distinguish probe activation, probe cleanup, caller-thread
application acquisition, and session acquisition, with start and completion
messages. The last stage helps locate a hang. A reused connection should
produce no new acquisition messages.

For automated coverage, manual regression procedures, implementation history,
and VM testing guidance, see [maintainer notes][verification].

[verification]: DevelopmentNotes.md#outlook-verification

## References

- [COM single-threaded apartments](https://learn.microsoft.com/en-us/windows/win32/com/single-threaded-apartments)
- [Thread.Join and its timeout range](https://learn.microsoft.com/en-us/dotnet/api/system.threading.thread.join)
- [Marshal.ReleaseComObject and shared RCW risks](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject)
