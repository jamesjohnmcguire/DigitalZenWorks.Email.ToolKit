/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceIntegrationTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.Diagnostics;
using System.Threading;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Tests the public service with Outlook installed and an interactive
/// user session. These integration tests run by default.
/// </summary>
[Category("Integration")]
[Apartment(ApartmentState.STA)]
[NonParallelizable]
internal sealed class OutlookServiceIntegrationTests
{
	private IOutlookService? service;

	/// <summary>
	/// One time set up method.
	/// </summary>
	[OneTimeSetUp]
	public void OneTimeSetUp()
	{
		bool isOutlookPreviouslyStarted = OutlookService.IsOutlookStarted();

		if (isOutlookPreviouslyStarted == false)
		{
			StartOutlookIfNotRunning();
		}

		service = new OutlookService();
	}

	/// <summary>
	/// One time tear down method.
	/// </summary>
	[OneTimeTearDown]
	public void OneTimeTearDown()
	{
		if (service != null)
		{
			// Detaching is safe even when Outlook predates this fixture.
			service.Disconnect();
		}
	}

	/// <summary>
	/// Verifies the public API connects and keeps a stable session.
	/// </summary>
	/// <remarks>Prerequisites: Outlook must be running in the current
	/// user session.  This test attempts to attach to an existing Outlook
	/// process and asserts that Connect returns true and a session is
	/// available.</remarks>
	[Test]
	public void AttachToExistingOutlookWhenOutlookRunningSuccess()
	{
		bool connected = service!.Connect(timeOutSeconds: 33);

		Assert.That(connected, Is.True);
		Assert.That(service.IsConnected, Is.True);
		Assert.That(service.Session, Is.Not.Null);
		IOutlookSession? first = service.Session;

		Assert.That(service.Connect(), Is.True);
		Assert.That(service.Session, Is.SameAs(first));
		Assert.That(OutlookService.IsOutlookStarted(), Is.True);
	}

	/// <summary>
	/// Verifies installation detection on the configured test machine.
	/// </summary>
	[Test]
	public void IsOutlookInstalledReportsAvailableInstallation()
	{
		Assert.That(OutlookService.IsOutlookInstalled(), Is.True);
	}

	/// <summary>
	/// Verifies disconnect leaves an independently held Outlook application
	/// usable and permits the service to connect again on the same STA.
	/// </summary>
	[Test]
	public void DisconnectPreservesExistingOutlookAndAllowsReconnect()
	{
		using OutlookTestContext context = new();
		OutlookService localService = new();

		try
		{
			Assert.That(localService.Connect(33), Is.True);
			IOutlookSession? first = localService.Session;

			localService.Disconnect();
			localService.Disconnect();

			Assert.That(localService.IsConnected, Is.False);
			Assert.That(localService.Session, Is.Null);

			// Process detection alone would miss a disconnected RCW.
			// Exercise the independently acquired application and namespace.
			Assert.That(context.Application.Version, Is.Not.Empty);
			Outlook.Stores stores = context.Track(context.NameSpace.Stores);
			Assert.That(stores.Count, Is.GreaterThan(0));

			Assert.That(localService.Connect(33), Is.True);
			Assert.That(localService.Session, Is.Not.Null);
			Assert.That(localService.Session, Is.Not.SameAs(first));
		}
		finally
		{
			localService.Disconnect();
		}
	}

	private static void StartOutlookIfNotRunning()
	{
		string outlookPath =
			@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE";

		Process.Start(outlookPath);
		Thread.Sleep(3000);
	}
}
