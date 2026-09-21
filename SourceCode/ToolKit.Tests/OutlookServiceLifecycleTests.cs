/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceLifecycleTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;

internal sealed class OutlookServiceLifecycleTests
{
	/// <summary>
	/// Tests that Connect fails when the factory reports Outlook as
	/// unavailable.
	/// </summary>
	[Test]
	public void ConnectFailWhenFactoryReportsUnavailable()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();

		factory.IsAvailable = false;
		factory.Connection = null;

		bool result = service.Connect(factory);

		Assert.IsFalse(result);
		Assert.AreEqual(1, factory.IsOutlookAvailableCallCount);
		Assert.AreEqual(0, factory.CreateConnectionCallCount);
		Assert.IsNull(service.Session);
	}

	/// <summary>
	/// Tests that Connect handles exceptions thrown by the factory and returns
	/// false.
	/// </summary>
	[Test]
	public void ConnectHandlesFactoryExceptionsReturnsFalse()
	{
		IOutlookFactory throwingFactory = new ThrowingFactory();

		OutlookService service = new();

		bool result = service.Connect(throwingFactory);

		Assert.That(result, Is.False);
		Assert.That(service.Session, Is.Null);
	}

	/// <summary>
	/// Tests that Connect is idempotent when already connected. Calling Connect
	/// multiple times should return the same result.
	/// </summary>
	[Test]
	public void ConnectIsIdempotentWhenAlreadyConnected()
	{
		OutlookService service = new();

		FakeOutlookSession session = new();
		FakeOutlookConnection connection1 = new(session);
		FakeOutlookFactory factory = new();

		factory.IsAvailable = true;
		factory.Connection = connection1;

		bool first = service.Connect(factory);
		bool second = service.Connect(factory);

		Assert.That(first, Is.True);
		Assert.That(second, Is.True);
		Assert.That(factory.CreateConnectionCallCount, Is.EqualTo(1));
	}

	/// <summary>
	/// Tests that Connect succeeds when the factory creates a connection and
	/// does not throw an exception. The session should be available after a
	/// successful connection.
	/// </summary>
	[Test]
	public void ConnectSuccessWhenFactoryCreatesConnection()
	{
		OutlookService service = new();

		FakeOutlookSession session = new();
		FakeOutlookConnection connection = new(session);
		FakeOutlookFactory factory = new();

		factory.IsAvailable = true;
		factory.Connection = connection;

		bool result = service.Connect(factory);

		Assert.That(result, Is.True);
		Assert.That(service.Session, Is.Not.Null);
		Assert.That(factory.CreateConnectionCallCount, Is.EqualTo(1));
	}

	/// <summary>
	/// Tests that Disconnect releases the session and allows a subsequent
	/// connection.
	/// </summary>
	[Test]
	public void DisconnectReleasesSessionAndAllowsReconnect()
	{
		OutlookService service = new();

		FakeOutlookSession session1 = new();
		FakeOutlookConnection connection1 = new(session1);
		FakeOutlookFactory factory = new();

		factory.IsAvailable = true;
		factory.Connection = connection1;

		bool connected = service.Connect(factory);

		Assert.That(connected, Is.True);
		Assert.That(service.Session, Is.Not.Null);

		service.Disconnect();

		// Allow reconnect with a new connection
		FakeOutlookSession session2 = new();
		FakeOutlookConnection connection2 = new(session2);
		factory.Connection = connection2;

		bool reconnected = service.Connect(factory);

		Assert.That(reconnected, Is.True);
		Assert.That(service.Session, Is.Not.Null);
	}
}
