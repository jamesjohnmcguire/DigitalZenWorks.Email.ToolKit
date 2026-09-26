/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceIntegrationTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;
using System.Diagnostics;
using System.Threading;

/// <summary>
/// Integration tests for OutlookService. These tests are explicit/manual and
/// should only be run on a developer machine with Outlook installed and an
/// interactive user session. They are excluded from CI by being marked
/// [Explicit] and [Category("Integration")].
/// </summary>
[Category("Integration")]
[Apartment(ApartmentState.STA)]
internal sealed class OutlookServiceIntegrationTests
{
	private bool isOutlookPreviouslyStarted;

	private OutlookService? service;

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

		service = new();
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
	/// Attempts to attach to an existing Outlook instance. Run manually.
	/// </summary>
	/// <remarks>Manual prerequisites: Outlook must be running in the current
	/// user session.  This test attempts to attach to an existing Outlook
	/// process and asserts that Connect returns true and a session is
	/// available.</remarks>
	[Test]
	public void AttachToExistingOutlookWhenOutlookRunningSuccess()
	{
		OutlookFactory factory = new();

		bool connected = service.Connect(factory, timeOutSeconds: 33);

		Assert.That(connected, Is.True);
		Assert.That(service.Session, Is.Not.Null);
	}

	private static void StartOutlookIfNotRunning()
	{
		string outlookPath =
			@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE";

		Process.Start(outlookPath);
		Thread.Sleep(3000);
	}
}
