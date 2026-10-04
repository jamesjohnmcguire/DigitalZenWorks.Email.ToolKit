/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookApplicationProbeTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using NUnit.Framework;

/// <summary>
/// Exercises real STA probe scheduling without activating Outlook.
/// Gates establish ordering; time limits only guard against stuck tests.
/// </summary>
internal sealed class OutlookApplicationProbeTests
{
	/// <summary>
	/// Verifies activation and cleanup stay in a temporary STA apartment.
	/// </summary>
	[Test]
	public void ProbeRunsOnSeparateStaThread()
	{
		Thread caller = Thread.CurrentThread;
		Thread? worker = null;
		ApartmentState apartment = ApartmentState.Unknown;
		OutlookApplicationProbe probe = new(() =>
		{
			worker = Thread.CurrentThread;
			apartment = worker.GetApartmentState();
		});

		bool available = probe.CanCreateApplication(5);

		Assert.That(available, Is.True);

		// Returning success requires the worker and its cleanup to finish.
		Assert.That(worker, Is.Not.Null);
		Assert.That(worker!.IsAlive, Is.False);

		// A caller must not receive a COM object from this short-lived STA.
		Assert.That(worker, Is.Not.SameAs(caller));
		Assert.That(apartment, Is.EqualTo(ApartmentState.STA));
	}

	/// <summary>
	/// Verifies invalid waits cannot leave an unobserved activation running.
	/// </summary>
	/// <param name="seconds">An unsupported probe wait.</param>
	[TestCase(-1)]
	[TestCase(2147484)]
	[TestCase(int.MaxValue)]
	public void InvalidWaitDoesNotStartProbe(int seconds)
	{
		int calls = 0;
		OutlookApplicationProbe probe = new(() =>
		{
			Interlocked.Increment(ref calls);
		});

		Assert.That(
			() => probe.CanCreateApplication(seconds),
			Throws.TypeOf<ArgumentOutOfRangeException>());
		Assert.That(calls, Is.Zero);
	}

	/// <summary>
	/// Verifies the largest whole-second wait supported by Thread.Join.
	/// </summary>
	[Test]
	public void MaximumWaitIsAccepted()
	{
		OutlookApplicationProbe probe = new(() => { });

		Assert.That(probe.CanCreateApplication(2147483), Is.True);
	}

	/// <summary>
	/// Verifies a completed failure does not poison subsequent attempts.
	/// </summary>
	/// <param name="comFailure">Whether activation throws a COM error.</param>
	[TestCase(true)]
	[TestCase(false)]
	[SuppressMessage(
		"Usage",
		"CA2201:Do not raise reserved exception types",
		Justification = "Simulate COM activation errors without using Outlook.")]
	public void CompletedFailureAllowsFreshProbe(bool comFailure)
	{
		int calls = 0;
		OutlookApplicationProbe probe = new(() =>
		{
			calls++;

			if (calls == 1)
			{
				if (comFailure)
				{
					throw new COMException("Activation failed");
				}

				throw new InvalidOperationException("Activation failed");
			}
		});

		Assert.That(probe.CanCreateApplication(5), Is.False);
		Assert.That(probe.CanCreateApplication(5), Is.True);

		// Availability is transient: neither failure nor success is cached.
		Assert.That(probe.CanCreateApplication(5), Is.True);
		Assert.That(calls, Is.EqualTo(3));
	}

	/// <summary>
	/// Verifies unexpected worker errors are transported to a waiting caller.
	/// </summary>
	[Test]
	public void UnexpectedFailureIsRethrownToWaitingCaller()
	{
		ArgumentException failure = new("Unexpected probe failure");
		OutlookApplicationProbe probe = new(() => throw failure);

		Assert.That(
			() => probe.CanCreateApplication(5),
			Throws.Exception.SameAs(failure));
	}

	/// <summary>
	/// Verifies retries through different factories share an outstanding
	/// probe, including cleanup after the original caller stopped waiting.
	/// </summary>
	/// <param name="failCleanup">Whether late cleanup should fail.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void TimedOutProbeRemainsActiveUntilCleanupFinishes(
		bool failCleanup)
	{
		using ManualResetEventSlim cleanupStarted = new(false);
		using ManualResetEventSlim allowCleanup = new(false);
		Thread? firstWorker = null;
		int calls = 0;
		OutlookApplicationProbe probe = new(() =>
		{
			int call = Interlocked.Increment(ref calls);

			if (call == 1)
			{
				firstWorker = Thread.CurrentThread;

				// This gate represents COM cleanup after activation returns.
				// The coordinator must wait for the entire operation to exit.
				cleanupStarted.Set();
				allowCleanup.Wait();

				if (failCleanup)
				{
					throw new InvalidOperationException("Late cleanup");
				}
			}
		});
		OutlookFactory firstFactory = new(probe);
		OutlookFactory retryFactory = new(probe);
		OutlookService service = new();

		try
		{
			// Zero is a supported nonblocking wait. The gate prevents a
			// scheduling race from turning this into a successful probe.
			Assert.That(service.Connect(firstFactory, 0), Is.False);
			Assert.That(cleanupStarted.Wait(TimeSpan.FromSeconds(5)), Is.True);

			Assert.That(retryFactory.CanCreateApplication(0), Is.False);
			Assert.That(firstFactory.CanCreateApplication(0), Is.False);

			// A new factory must not create another worker while cleanup runs.
			Assert.That(Volatile.Read(ref calls), Is.EqualTo(1));
			Assert.That(firstWorker!.IsAlive, Is.True);
		}
		finally
		{
			allowCleanup.Set();

			// Drain the worker before disposing the gates, even on failure.
			bool started = cleanupStarted.Wait(TimeSpan.FromSeconds(5));
			Assert.That(started, Is.True);
			Assert.That(firstWorker!.Join(TimeSpan.FromSeconds(5)), Is.True);
		}

		// Late completion must not publish a connection into a timed-out
		// service. It only makes room for a fresh availability probe.
		Assert.That(service.IsConnected, Is.False);
		Assert.That(service.Session, Is.Null);
		Assert.That(retryFactory.CanCreateApplication(5), Is.True);
		Assert.That(calls, Is.EqualTo(2));
	}
}
