/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceFailureTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Runtime.InteropServices;
using NUnit.Framework;

/// <summary>
/// Tests connection failures, retries, and the public service state.
/// </summary>
internal sealed class OutlookServiceFailureTests
{
	/// <summary>
	/// Verifies a new service exposes no connection or session.
	/// </summary>
	[Test]
	public void NewServiceIsDisconnected()
	{
		OutlookService service = new();

		Assert.That(service.IsConnected, Is.False);
		Assert.That(service.Session, Is.Null);
		service.Disconnect();
		Assert.That(service.IsConnected, Is.False);
		Assert.That(service.Session, Is.Null);
	}

	/// <summary>
	/// Verifies explicit timeouts are forwarded without changing the value.
	/// </summary>
	/// <param name="seconds">The requested timeout.</param>
	[TestCase(0)]
	[TestCase(1)]
	[TestCase(33)]
	public void ConnectForwardsTimeout(int seconds)
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();

		Assert.That(service.Connect(factory, seconds), Is.False);
		Assert.That(factory.LastTimeoutSeconds, Is.EqualTo(seconds));
		Assert.That(factory.CreateConnectionCallCount, Is.Zero);
		Assert.That(service.IsConnected, Is.False);
	}

	/// <summary>
	/// Verifies the default availability timeout.
	/// </summary>
	[Test]
	public void ConnectUsesDefaultTimeout()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();

		service.Connect(factory);

		Assert.That(factory.LastTimeoutSeconds, Is.EqualTo(10));
	}

	/// <summary>
	/// Verifies unsuccessful attempts leave the service ready to retry.
	/// </summary>
	/// <param name="stage">
	/// 0: Outlook unavailable; 1: null connection; 2: null session.
	/// </param>
	[TestCase(0)]
	[TestCase(1)]
	[TestCase(2)]
	public void ConnectRetriesAfterUnsuccessfulAttempt(int stage)
	{
		// The fakes select where the first attempt fails without starting
		// Outlook. The same service is then used for a successful retry.
		OutlookService service = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = false;
		factory.Connection = null;

		// These are cumulative CreateConnection counts after the first
		// attempt and after the retry. Stages 1 and 2 reach creation twice.
		int expectedCallCount = 1;
		int expectedCallCount2 = 2;

		// Stages 1 and 2 pass availability so the attempt can fail later.
		if (stage != 0)
		{
			factory.IsAvailable = true;
		}

		// Stage 2 supplies a connection whose session is explicitly null.
		// A connection alone is not sufficient for this test's success
		// contract: the service must also acquire a usable session.
		if (stage == 2)
		{
			factory.Connection = new FakeOutlookConnection(null);
		}

		// Stage 0 stops at availability, so only the retry should create
		// a connection.
		if (stage == 0)
		{
			expectedCallCount = 0;
			expectedCallCount2 = 1;
		}

		bool result = service.Connect(factory);

		// Each configured failure must be reported to the caller. This
		// includes stage 2, where a non-null connection has no session.
		Assert.That(result, Is.False);

		// Failed attempts must leave the service disconnected. Retaining
		// a partial connection could make the retry skip the factory.
		Assert.That(service.IsConnected, Is.False);

		// Callers must not receive a session from an unsuccessful attempt.
		// Check the exposed session separately from the connection flag.
		Assert.That(service.Session, Is.Null);

		// Availability failure must prevent connection creation (stage 0).
		// Later failures must reach creation exactly once (stages 1 and 2),
		// showing that the intended failure path was exercised.
		Assert.That(
			factory.CreateConnectionCallCount,
			Is.EqualTo(expectedCallCount));

		// Remove the failure using the same factory and service. A fresh
		// service would not demonstrate recovery from the earlier attempt.
		FakeOutlookSession expectedSession = new();
		factory.IsAvailable = true;
		factory.Connection = new FakeOutlookConnection(expectedSession);

		bool retryResult = service.Connect(factory);

		// Once a usable connection and session are available, the caller
		// must be able to retry successfully without recreating the service.
		Assert.That(retryResult, Is.True);

		// Success must also be reflected in the service's public state.
		Assert.That(service.IsConnected, Is.True);

		// Identity proves that the service exposes the retry's session,
		// rather than retaining null or substituting another session.
		Assert.That(service.Session, Is.SameAs(expectedSession));

		// Availability must be checked on both attempts so an earlier
		// failure does not prevent the service noticing recovery.
		Assert.That(factory.IsOutlookAvailableCallCount, Is.EqualTo(2));

		// The retry must make one additional creation call. The total is
		// one for stage 0 and two for stages 1 and 2. This guards against
		// reusing a failed connection or creating unnecessary connections.
		Assert.That(
			factory.CreateConnectionCallCount,
			Is.EqualTo(expectedCallCount2));
	}

	/// <summary>
	/// Verifies handled failures at each connection stage permit retry.
	/// </summary>
	/// <param name="stage">Availability, creation, or session acquisition.</param>
	/// <param name="comFailure">Whether to inject a COM failure.</param>
	[System.Diagnostics.CodeAnalysis.SuppressMessage(
		"Usage",
		"CA2201:Do not raise reserved exception types",
		Justification = "Inject the COM failure handled by the service.")]
	[TestCase(0, false)]
	[TestCase(0, true)]
	[TestCase(1, false)]
	[TestCase(1, true)]
	[TestCase(2, false)]
	[TestCase(2, true)]
	public void ConnectHandlesFailureAndRetries(int stage, bool comFailure)
	{
		Exception failure = comFailure ?
			new COMException("Outlook unavailable") :
			new InvalidOperationException("Outlook unavailable");
		OutlookService service = new();
		FakeOutlookConnection connection = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;

		if (stage == 0)
		{
			factory.AvailabilityException = failure;
		}
		else if (stage == 1)
		{
			factory.ConnectionException = failure;
		}
		else
		{
			connection.SessionException = failure;
		}

		Assert.That(service.Connect(factory), Is.False);
		Assert.That(service.IsConnected, Is.False);
		Assert.That(service.Session, Is.Null);
		Assert.That(
			factory.CreateConnectionCallCount,
			Is.EqualTo(stage == 0 ? 0 : 1));
		Assert.That(
			connection.SessionAccessCount, Is.EqualTo(stage == 2 ? 1 : 0));

		FakeOutlookSession expectedSession = new();
		factory.AvailabilityException = null;
		factory.ConnectionException = null;
		factory.Connection = new FakeOutlookConnection(expectedSession);

		Assert.That(service.Connect(factory), Is.True);
		Assert.That(service.IsConnected, Is.True);
		Assert.That(service.Session, Is.SameAs(expectedSession));
		Assert.That(factory.IsOutlookAvailableCallCount, Is.EqualTo(2));
		Assert.That(
			factory.CreateConnectionCallCount,
			Is.EqualTo(stage == 0 ? 1 : 2));
	}

	/// <summary>
	/// Verifies unrelated errors are not hidden as Outlook unavailability.
	/// </summary>
	[Test]
	public void ConnectPropagatesUnexpectedFactoryFailure()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();
		ArgumentException failure = new("Unexpected factory error");
		factory.AvailabilityException = failure;

		Assert.That(
			() => service.Connect(factory), Throws.Exception.SameAs(failure));
		Assert.That(factory.CreateConnectionCallCount, Is.Zero);
		Assert.That(service.IsConnected, Is.False);
		Assert.That(service.Session, Is.Null);
	}
}
