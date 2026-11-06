/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookSessionFactoryTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using NUnit.Framework;

/// <summary>
/// Tests helper factories without activating Outlook or creating COM fakes.
/// </summary>
internal sealed class OutlookSessionFactoryTests
{
	/// <summary>
	/// Verifies folder helpers accept the concrete session implementation
	/// through its interface without requiring Outlook activation.
	/// </summary>
	[Test]
	public void FolderFromSessionAcceptsConcreteSession()
	{
		IOutlookSession session = new OutlookSession(null);

		OutlookFolder folder = OutlookFolder.FromSession(session);

		Assert.That(folder, Is.Not.Null);
	}

	/// <summary>
	/// Verifies the folder factory rejects unsupported session references.
	/// </summary>
	/// <param name="useNull">Whether to supply null rather than a different
	/// implementation of the session interface.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void FolderFromSessionRejectsUnsupportedSession(bool useNull)
	{
		IOutlookSession? session = null;

		if (useNull == false)
		{
			session = new FakeOutlookSession();
		}

		Assert.That(
			() => OutlookFolder.FromSession(session!),
			Throws.TypeOf<NotSupportedException>());
	}

	/// <summary>
	/// Verifies store helpers retain the supplied concrete session and use
	/// its disconnected behavior without starting Outlook.
	/// </summary>
	[Test]
	public void StoreFromSessionAcceptsConcreteSession()
	{
		OutlookSession session = new(null);

		OutlookStore store = OutlookStore.FromSession(session);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(store.OutlookSession, Is.SameAs(session));
			Assert.That(store.GetItemFromEntryId("entry"), Is.Null);
		}
	}

	/// <summary>
	/// Verifies the store factory rejects unsupported session references.
	/// </summary>
	/// <param name="useNull">Whether to supply null rather than a different
	/// implementation of the session interface.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void StoreFromSessionRejectsUnsupportedSession(bool useNull)
	{
		IOutlookSession? session = null;

		if (useNull == false)
		{
			session = new FakeOutlookSession();
		}

		Assert.That(
			() => OutlookStore.FromSession(session!),
			Throws.TypeOf<NotSupportedException>());
	}
}
