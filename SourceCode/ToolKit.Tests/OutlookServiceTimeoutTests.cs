/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceTimeoutTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;

/// <summary>
/// Specifies a deadline for obtaining a complete Outlook connection.
/// Gates simulate blocked dependencies without creating Outlook COM objects.
/// </summary>
[Apartment(ApartmentState.STA)]
internal sealed class OutlookServiceTimeoutTests
{
	private static readonly TimeSpan Watchdog = TimeSpan.FromSeconds(5);

	/// <summary>
	/// Verifies cleanup of an unusable connection cannot delay the caller.
	/// </summary>
	[Test]
	public void ConnectDoesNotWaitForBlockedCleanup()
	{
		OutlookService service = new();
		BlockingOperation cleanup = new();
		FakeOutlookConnection connection = new(null);
		connection.BeforeDispose = cleanup.Invoke;
		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;

		try
		{
			bool connected = service.Connect(factory, 1);

			Assert.That(connected, Is.False);
			Assert.That(
				cleanup.WaitForStart(),
				Is.True,
				"An unusable connection must be cleaned up.");
			Assert.That(
				cleanup.IsWaiting,
				Is.True,
				"Connect must return while cleanup is still blocked.");
			Assert.That(service.IsConnected, Is.False);
			Assert.That(service.Session, Is.Null);
			Assert.That(connection.QuitCallCount, Is.Zero);
		}
		finally
		{
			cleanup.Release();
			cleanup.WaitForCompletion();
			service.Disconnect();
		}
	}

	/// <summary>
	/// Verifies a successful result arriving after timeout is discarded.
	/// </summary>
	/// <param name="blockSession">Whether to block session acquisition rather
	/// than connection creation.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void ConnectIgnoresLateSuccessAfterTimeout(bool blockSession)
	{
		OutlookService service = new();
		BlockingOperation acquisition = new();
		TaskCompletionSource<bool> disposed = new(
			TaskCreationOptions.RunContinuationsAsynchronously);
		FakeOutlookConnection connection = new();
		connection.BeforeDispose = () =>
		{
			disposed.TrySetResult(true);
		};
		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;

		if (blockSession == true)
		{
			connection.BeforeGetSession = acquisition.Invoke;
		}
		else
		{
			factory.BeforeCreateConnection = acquisition.Invoke;
		}

		try
		{
			bool connected = service.Connect(factory, 1);

			Assert.That(connected, Is.False);
			Assert.That(acquisition.IsWaiting, Is.True);
			Assert.That(service.IsConnected, Is.False);
			Assert.That(service.Session, Is.Null);

			// Let the expired attempt produce its successful result. Observe
			// production cleanup before checking that it was never published.
			acquisition.Release();
			Assert.That(disposed.Task.Wait(Watchdog), Is.True);
			Assert.That(service.IsConnected, Is.False);
			Assert.That(service.Session, Is.Null);
			Assert.That(connection.DisposeCallCount, Is.EqualTo(1));
			Assert.That(connection.QuitCallCount, Is.Zero);
		}
		finally
		{
			acquisition.Release();
			acquisition.WaitForCompletion();
			service.Disconnect();
		}
	}

	/// <summary>
	/// Verifies sequential retries cannot accumulate blocked acquisitions.
	/// </summary>
	[Test]
	public void ConnectRetriesDoNotStartAnotherBlockedAcquisition()
	{
		OutlookService service = new();
		BlockingOperation acquisition = new();
		FakeOutlookConnection connection = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;
		int acquisitionCount = 0;
		factory.BeforeCreateConnection = () =>
		{
			Interlocked.Increment(ref acquisitionCount);
			acquisition.Invoke();
		};

		try
		{
			bool first = service.Connect(factory, 1);

			Assert.That(first, Is.False);
			Assert.That(acquisition.IsWaiting, Is.True);

			// Both calls use the same caller thread. The first Connect has
			// returned, but its underlying acquisition is still blocked.
			bool second = service.Connect(factory, 1);

			Assert.That(second, Is.False);
			Assert.That(acquisition.IsWaiting, Is.True);
			Assert.That(Volatile.Read(ref acquisitionCount), Is.EqualTo(1));
			Assert.That(service.IsConnected, Is.False);
			Assert.That(service.Session, Is.Null);
		}
		finally
		{
			acquisition.Release();
			acquisition.WaitForCompletion();
			service.Disconnect();
		}
	}

	/// <summary>
	/// Verifies the deadline includes every dependency acquisition stage.
	/// </summary>
	/// <param name="stage">Availability, creation, or session
	/// acquisition.</param>
	[TestCase(0)]
	[TestCase(1)]
	[TestCase(2)]
	public void ConnectReturnsFalseWhenAcquisitionIsBlocked(int stage)
	{
		OutlookService service = new();
		BlockingOperation acquisition = new();
		FakeOutlookConnection connection = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;

		if (stage == 0)
		{
			factory.BeforeAvailabilityCheck = acquisition.Invoke;
		}
		else if (stage == 1)
		{
			factory.BeforeCreateConnection = acquisition.Invoke;
		}
		else
		{
			connection.BeforeGetSession = acquisition.Invoke;
		}

		try
		{
			Stopwatch elapsed = Stopwatch.StartNew();
			bool connected = service.Connect(factory, 1);
			elapsed.Stop();

			Assert.That(connected, Is.False);
			Assert.That(
				acquisition.IsWaiting,
				Is.True,
				"The simulated operation must remain blocked after timeout.");
			Assert.That(service.IsConnected, Is.False);
			Assert.That(service.Session, Is.Null);

			// Allow scheduling overhead without accepting an immediate failure
			// or waiting for the five-second gate watchdog to release the call.
			Assert.That(
				elapsed.ElapsedMilliseconds, Is.GreaterThanOrEqualTo(900));
			Assert.That(elapsed.Elapsed, Is.LessThan(TimeSpan.FromSeconds(3)));
		}
		finally
		{
			acquisition.Release();
			acquisition.WaitForCompletion();
			service.Disconnect();
		}
	}

	/// <summary>
	/// Verifies completed stages consume the next stage's remaining budget.
	/// </summary>
	/// <param name="includeSession">Whether to delay creation and session
	/// acquisition rather than availability and creation.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void ConnectUsesOneTimeoutAcrossAcquisitionStages(
		bool includeSession)
	{
		OutlookService service = new();
		BlockingOperation firstStage = new();
		BlockingOperation secondStage = new();
		FakeOutlookConnection connection = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;
		Action firstDelay = () =>
		{
			firstStage.InvokeFor(TimeSpan.FromSeconds(2));
		};
		Action secondDelay = () =>
		{
			secondStage.InvokeFor(TimeSpan.FromSeconds(2));
		};

		if (includeSession == true)
		{
			factory.BeforeCreateConnection = firstDelay;
			connection.BeforeGetSession = secondDelay;
		}
		else
		{
			factory.BeforeAvailabilityCheck = firstDelay;
			factory.BeforeCreateConnection = secondDelay;
		}

		try
		{
			// Each stage fits within three seconds, but their combined four
			// seconds must exceed the single budget for the whole connection.
			bool connected = service.Connect(factory, 3);

			Assert.That(connected, Is.False);
			Assert.That(firstStage.HasCompleted, Is.True);
			Assert.That(secondStage.IsWaiting, Is.True);
			Assert.That(service.IsConnected, Is.False);
			Assert.That(service.Session, Is.Null);
		}
		finally
		{
			firstStage.Release();
			secondStage.Release();
			firstStage.WaitForCompletion();
			secondStage.WaitForCompletion();
			service.Disconnect();
		}
	}

	/// <summary>
	/// Holds a dependency until released, with a watchdog for failing tests.
	/// Task gates also remain safe if an abandoned worker starts late.
	/// </summary>
	private sealed class BlockingOperation
	{
		private readonly TaskCompletionSource<bool> started = new(
			TaskCreationOptions.RunContinuationsAsynchronously);

		private readonly TaskCompletionSource<bool> release = new(
			TaskCreationOptions.RunContinuationsAsynchronously);

		private readonly TaskCompletionSource<bool> completed = new(
			TaskCreationOptions.RunContinuationsAsynchronously);

		/// <summary>
		/// Gets a value indicating whether the simulated operation finished.
		/// </summary>
		public bool HasCompleted => completed.Task.IsCompleted;

		/// <summary>
		/// Gets a value indicating whether the simulated operation is blocked.
		/// </summary>
		public bool IsWaiting =>
			started.Task.IsCompleted && completed.Task.IsCompleted == false;

		/// <summary>
		/// Blocks until released, or until the safety watchdog expires.
		/// </summary>
		public void Invoke()
		{
			InvokeFor(Watchdog);
		}

		/// <summary>
		/// Blocks for a controlled stage duration unless released earlier.
		/// </summary>
		/// <param name="duration">The stage duration or safety
		/// watchdog.</param>
		public void InvokeFor(TimeSpan duration)
		{
			using Timer timer = new(
				_ => Release(), null, duration, Timeout.InfiniteTimeSpan);
			started.TrySetResult(true);

			try
			{
				release.Task.GetAwaiter().GetResult();
			}
			finally
			{
				completed.TrySetResult(true);
			}
		}

		/// <summary>
		/// Allows all invocations to finish, including ones that start later.
		/// </summary>
		public void Release()
		{
			release.TrySetResult(true);
		}

		/// <summary>
		/// Drains an operation that entered its blocking callback.
		/// </summary>
		public void WaitForCompletion()
		{
			if (started.Task.IsCompleted == true)
			{
				Assert.That(
					completed.Task.Wait(Watchdog),
					Is.True,
					"The simulated dependency did not finish after release.");
			}
		}

		/// <summary>
		/// Waits for asynchronous cleanup to enter its blocking callback.
		/// </summary>
		/// <returns>Whether the callback started before the watchdog.</returns>
		public bool WaitForStart()
		{
			bool hasStarted = started.Task.Wait(Watchdog);
			return hasStarted;
		}
	}
}
