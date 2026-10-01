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

/// <summary>
/// Tests the public service with Outlook installed and an interactive
/// user session. These integration tests run by default.
/// </summary>
[Category("Integration")]
[Apartment(ApartmentState.STA)]
[NonParallelizable]
internal sealed class OutlookServiceIntegrationTests
{
	private bool isOutlookPreviouslyStarted;

	private IOutlookService? service;

	/// <summary>
	/// One time set up method.
	/// </summary>
	[OneTimeSetUp]
	public void OneTimeSetUp()
	{
		isOutlookPreviouslyStarted = OutlookService.IsOutlookStarted();

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
		if (service != null && isOutlookPreviouslyStarted == false)
		{
			// Clean up
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

	private static void StartOutlookIfNotRunning()
	{
		string outlookPath =
			@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE";

		Process.Start(outlookPath);
		Thread.Sleep(3000);
	}
}
