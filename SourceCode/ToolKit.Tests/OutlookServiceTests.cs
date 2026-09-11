/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;

internal sealed class OutlookServiceTests
{
	[Test]
	public void Connect_CallsFactory_WhenNoExistingOutlook()
	{
		var service = new OutlookService();

		var factory = new FakeOutlookFactory
		{
			IsAvailable = false
		};

		service.Connect(factory);

		Assert.That(factory.CreateConnectionCallCount, Is.EqualTo(0));

		Assert.That(
			factory.IsOutlookAvailableCallCount,
			Is.EqualTo(1));
	}

	[Test]
	public void Connect_ChecksAvailability()
	{
		OutlookService service = new();

		FakeOutlookFactory factory = new()
		{
			IsAvailable = false
		};

		service.Connect(factory);

		Assert.That(
			factory.IsOutlookAvailableCallCount,
			Is.EqualTo(1));
	}

	[Test]
	public void Connect_DoesNotCreateConnection_WhenUnavailable()
	{
		OutlookService service = new();

		FakeOutlookFactory factory = new()
		{
			IsAvailable = false
		};

		service.Connect(factory);

		Assert.That(
			factory.CreateConnectionCallCount,
			Is.EqualTo(0));
	}

	[Test]
	public void Connect_DoesNotReconnect_WhenAlreadyConnected()
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

	[Test]
	public void Connect_IgnoresFactory_AfterAlreadyConnected()
	{
		OutlookService service = new();

		FakeOutlookSession session = new();

		FakeOutlookConnection connection =
			new(session);

		FakeOutlookFactory firstFactory = new()
		{
			IsAvailable = true,
			Connection = connection
		};

		FakeOutlookFactory secondFactory = new()
		{
			IsAvailable = false
		};

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

	[Test]
	public void Connect_ReturnsFalse_WhenOutlookUnavailable()
	{
		OutlookService service = new();

		FakeOutlookFactory factory = new()
		{
			IsAvailable = false
		};

		bool connected = service.Connect(factory);

		Assert.That(connected, Is.False);
		Assert.That(service.Session, Is.Null);
	}

	[Test]
	public void Connect_ReturnsTrue_WhenConnectionCreated()
	{
		OutlookService service = new();

		FakeOutlookSession session = new();

		FakeOutlookConnection connection = new(session);

		FakeOutlookFactory factory = new()
		{
			IsAvailable = true,
			Connection = connection
		};

		bool result = service.Connect(factory);

		Assert.That(result, Is.True);
	}

	[Test]
	public void Connect_SetsSession_WhenConnectionCreated()
	{
		OutlookService service = new();

		FakeOutlookSession expectedSession = new();

		FakeOutlookConnection connection =
			new(expectedSession);

		FakeOutlookFactory factory = new()
		{
			IsAvailable = true,
			Connection = connection
		};

		bool result = service.Connect(factory);

		Assert.That(result, Is.True);
		Assert.That(
			service.Session,
			Is.SameAs(expectedSession));
	}

	[Test]
	public void Session_IsNull_AfterFailedConnect()
	{
		var service = new OutlookService();

		var factory = new FakeOutlookFactory
		{
			IsAvailable = false
		};

		service.Connect(factory);

		Assert.That(service.Session, Is.Null);
	}

	[Test]
	public void Disconnect_BeforeConnect_DoesNotThrow()
	{
		var service = new OutlookService();

		Assert.DoesNotThrow(() =>
		{
			service.Disconnect();
		});
	}
}
