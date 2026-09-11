/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;

/// <summary>
/// Tests for the <see cref="OutlookService"/> class.
/// </summary>
internal sealed class OutlookServiceTests
{
	/// <summary>
	/// Tests that the Connect method checks availability when Outlook is not
	/// already connected.
	/// </summary>
	[Test]
	public void ConnectChecksAvailabilityWhenOutlookNotAlreadyConnected()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = false;

		service.Connect(factory);

		Assert.That(factory.CreateConnectionCallCount, Is.EqualTo(0));

		// Ensures Connect actually checks availability rather than
		// simply returning false.
		Assert.That(
			factory.IsOutlookAvailableCallCount,
			Is.EqualTo(1));
	}

	/// <summary>
	/// Tests that the Connect method checks availability when Outlook is
	/// already connected.
	/// </summary>
	[Test]
	public void ConnectChecksAvailability()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = false;

		service.Connect(factory);

		Assert.That(
			factory.IsOutlookAvailableCallCount,
			Is.EqualTo(1));
	}

	/// <summary>
	/// Tests that the Connect method does not create a connection when Outlook
	/// is unavailable.
	/// </summary>
	[Test]
	public void ConnectDoesNotCreateConnectionWhenUnavailable()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = false;

		service.Connect(factory);

		Assert.That(
			factory.CreateConnectionCallCount,
			Is.EqualTo(0));
	}

	/// <summary>
	/// Tests that the Connect method does not create a new connection when
	/// Outlook is already connected.
	/// </summary>
	[Test]
	public void ConnectDoesNotReconnectWhenAlreadyConnected()
	{
		OutlookService service = new();
		FakeOutlookSession session = new();
		FakeOutlookConnection connection = new(session);

		FakeOutlookFactory factory = new();

		factory.IsAvailable = true;
		factory.Connection = connection;

		Assert.That(service.Connect(factory), Is.True);

		Assert.That(
			factory.CreateConnectionCallCount,
			Is.EqualTo(1));
	}

	/// <summary>
	/// Tests that the Connect method ignores the factory when Outlook is
	/// already connected.
	/// </summary>
	[Test]
	public void Connect_IgnoresFactory_AfterAlreadyConnected()
	{
		OutlookService service = new();
		FakeOutlookSession session = new();
		FakeOutlookConnection connection = new(session);

		FakeOutlookFactory firstFactory = new();
		firstFactory.IsAvailable = true;
		firstFactory.Connection = connection;

		FakeOutlookFactory secondFactory = new();
		secondFactory.IsAvailable = false;

		Assert.That(
			service.Connect(firstFactory),
			Is.True);

		Assert.That(
			service.Connect(secondFactory),
			Is.True);

		Assert.That(
			secondFactory.CreateConnectionCallCount,
			Is.EqualTo(0));
	}

	/// <summary>
	/// Tests that the Connect method returns false when Outlook is unavailable.
	/// </summary>
	[Test]
	public void ConnectReturnsFalseWhenOutlookUnavailable()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = false;

		bool connected = service.Connect(factory);

		Assert.That(connected, Is.False);
		Assert.That(service.Session, Is.Null);
	}

	/// <summary>
	/// Tests that the Connect method returns true when a connection is created.
	/// </summary>
	[Test]
	public void ConnectReturnsTrueWhenConnectionCreated()
	{
		OutlookService service = new();
		FakeOutlookSession session = new();
		FakeOutlookConnection connection = new(session);

		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;

		bool result = service.Connect(factory);

		Assert.That(result, Is.True);
	}

	/// <summary>
	/// Tests that the Connect method sets the Session property when a
	/// connection is created.
	/// </summary>
	[Test]
	public void ConnectSetsSessionWhenConnectionCreated()
	{
		OutlookService service = new();
		FakeOutlookSession expectedSession = new();
		FakeOutlookConnection connection = new(expectedSession);

		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;

		bool result = service.Connect(factory);

		Assert.That(result, Is.True);
		Assert.That(
			service.Session,
			Is.SameAs(expectedSession));
	}

	/// <summary>
	/// Tests that the Session property is null after a failed connection
	/// attempt.
	/// </summary>
	[Test]
	public void SessionIsNullAfterFailedConnect()
	{
		OutlookService service = new();
		FakeOutlookFactory factory = new();
		factory.IsAvailable = false;

		service.Connect(factory);

		Assert.That(service.Session, Is.Null);
	}

	/// <summary>
	/// Tests that calling Disconnect before Connect does not throw
	/// an exception.
	/// </summary>
	[Test]
	public void DisconnectBeforeConnectDoesNotThrow()
	{
		OutlookService service = new();

		Assert.DoesNotThrow(() =>
		{
			service.Disconnect();
		});
	}
}
