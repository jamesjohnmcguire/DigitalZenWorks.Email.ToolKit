# Outlook connection lifetime and probe timeout

This is the interim design for `OutlookService`, `OutlookFactory`, and
`OutlookConnection`. It isolates Outlook COM access while preserving the
existing public interfaces. It does not replace the legacy `OutlookAccount`
singleton throughout the library.

## Why retain this interim design

The timeout request became more important when a running but unresponsive
Outlook left activation blocked. Process detection could not establish
whether Outlook was usable. The temporary STA probe bounds the caller's
wait for that activation without returning a wrapper from a dying apartment.

For this branch, acceptance means failed probes publish no connection,
retries cannot accumulate outstanding probe workers, successful connections
are reused, and disconnect does not shut down shared Outlook. The second,
caller-thread acquisition remains unbounded. This is a deliberate interim
limit pending the larger apartment-ownership redesign.

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

## Verification and deferred work

Unit tests use the real coordinator with controlled operations and gates,
without manufacturing Outlook COM objects. They cover STA execution,
timeout validation, a blocked operation with a positive timeout, concurrent
factory callers sharing an outstanding probe, retries during cleanup, late
failure, and fresh probes after completion. Service tests cover reuse, disposal, failure cleanup,
reconnection, and compatibility with non-disposable connections.

Tests that activate Outlook or verify that an independent COM reference
survives disconnection are marked `Integration`. They require Outlook and
a configured interactive user session.

A strict overall deadline requires a larger design: a long-lived STA owner,
dispatch of all COM work to that apartment, and explicit rules for timed-out
work and cleanup. A deadline still would not make a hung COM call safely
cancellable. That redesign, deterministic lifetime management of all
borrowed wrappers, and any verified exclusive-shutdown policy are deferred.

## Stage logging and manual regression checks

Debug logs distinguish probe activation, probe cleanup, caller-thread
application acquisition, and session acquisition, with start and completion
messages. The last stage helps locate a hang. A reused connection should
produce no new acquisition messages.

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

## Optional VM testing

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

## References

- [COM single-threaded apartments](https://learn.microsoft.com/en-us/windows/win32/com/single-threaded-apartments)
- [Thread.Join and its timeout range](https://learn.microsoft.com/en-us/dotnet/api/system.threading.thread.join)
- [Marshal.ReleaseComObject and shared RCW risks](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject)
