/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookServiceIntegrationTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.Threading;
using NUnit.Framework;

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

		OutlookService service = new();

		bool connected = service.Connect(factory, timeOutSeconds: 3);

		Assert.That(connected, Is.True);
		Assert.That(service.Session, Is.Not.Null);

		// Clean up
		service.Disconnect();
	}
}
