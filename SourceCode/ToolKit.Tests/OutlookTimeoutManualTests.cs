/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookTimeoutManualTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Diagnostics;
using System.Threading;
using NUnit.Framework;

/// <summary>
/// Real Outlook scenarios that require an independently prepared host state.
/// Run each test in a separate process with an external watchdog.
/// </summary>
[Category("ManualOutlook")]
[Apartment(ApartmentState.STA)]
[NonParallelizable]
internal sealed class OutlookTimeoutManualTests
{
	/// <summary>
	/// Start with Outlook closed and a working profile. No other test should
	/// activate Outlook first. See Documentation/OutlookLifecycle.md.
	/// </summary>
	[Test]
	[Explicit("Run alone with Outlook closed and a configured profile.")]
	public void ConnectToClosedOutlookCreatesSession()
	{
		OutlookService service = new();

		try
		{
			Assert.That(service.Connect(10), Is.True);
			Assert.That(service.IsConnected, Is.True);
			Assert.That(service.Session, Is.Not.Null);
		}
		finally
		{
			service.Disconnect();
		}
	}

	/// <summary>
	/// Reproduce the original unresponsive Outlook state before starting this
	/// test. A stopped wait does not cancel an Outlook COM call, so use an
	/// external watchdog and record the last logged stage.
	/// </summary>
	[Test]
	[Explicit("Run alone after reproducing the original Outlook hang.")]
	public void UnresponsiveOutlookReturnsWithoutPublishingSession()
	{
		OutlookService service = new();
		Stopwatch stopwatch = Stopwatch.StartNew();

		try
		{
			bool connected = service.Connect(1);
			stopwatch.Stop();

			Assert.That(connected, Is.False);
			Assert.That(service.IsConnected, Is.False);
			Assert.That(service.Session, Is.Null);
			Assert.That(
				stopwatch.Elapsed, Is.LessThan(TimeSpan.FromSeconds(5)));
		}
		finally
		{
			service.Disconnect();
		}
	}
}
