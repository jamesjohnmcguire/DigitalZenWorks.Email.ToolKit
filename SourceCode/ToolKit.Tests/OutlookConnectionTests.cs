/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookConnectionTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;

/// <summary>
/// Tests stable session access through the connection wrapper.
/// </summary>
internal sealed class OutlookConnectionTests
{
	/// <summary>
	/// Verifies that calling Quit on an OutlookConnection does not throw when
	/// no Outlook application instance is present.
	/// </summary>
	/// <remarks>Uses a FakeOutlookSession and asserts no exception is thrown
	/// via Assert.DoesNotThrow.</remarks>
	[Test]
	public void QuitDoesNotThrowWhenNoApplication()
	{
		// Use internal test constructor that accepts a session.
		FakeOutlookSession session = new();
		using OutlookConnection connection = new(session);

		Assert.DoesNotThrow(() => connection.Quit());
	}

	/// <summary>
	/// Verifies that OutlookConnection.Session returns the same session
	/// instance provided to the constructor.
	/// </summary>
	[Test]
	public void SessionReturnsProvidedSessionWhenConstructedWithSession()
	{
		FakeOutlookSession session = new();
		using OutlookConnection connection = new(session);

		OutlookConnection abstraction = connection;
		IOutlookSession? first = abstraction.Session;
		IOutlookSession? second = abstraction.Session;

		Assert.That(first, Is.SameAs(session));
		Assert.That(second, Is.SameAs(first));
	}

	/// <summary>
	/// Verifies disposal relinquishes the connection's session and repeated
	/// cleanup is harmless. Borrowed session wrappers are not force-released.
	/// </summary>
	[Test]
	public void DisposeClearsSessionAndIsIdempotent()
	{
		FakeOutlookSession session = new();
		using OutlookConnection connection = new(session);
		IOutlookSession? first = connection.Session;

		connection.Dispose();
		connection.Dispose();

		Assert.That(connection.Session, Is.Null);
		Assert.That(first, Is.SameAs(session));
		Assert.DoesNotThrow(() => connection.Quit());
	}
}
