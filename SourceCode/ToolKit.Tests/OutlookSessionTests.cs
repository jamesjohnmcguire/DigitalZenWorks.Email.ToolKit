/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookSessionTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;

internal sealed class OutlookSessionTests
{
	/// <summary>
	/// Verifies GetItemFromId returns null when the OutlookSession has a null
	/// underlying session.
	/// </summary>
	/// <remarks>Initializes an OutlookSession with null, calls GetItemFromId
	/// using a non-existent id, and asserts that the result is null.</remarks>
	[Test]
	public void GetItemFromIdReturnsNullWhenSessionIsNull()
	{
		OutlookSession session = new(null);

		object? item = session.GetItemFromId("nonexistent");

		Assert.That(item, Is.Null);
	}

	/// <summary>
	/// Verifies that OutlookSession.OpenSharedItem returns null when the
	/// underlying session is null.
	/// </summary>
	/// <remarks>Creates an OutlookSession with a null underlying session and
	/// attempts to open a shared item, expecting a null result.</remarks>
	[Test]
	public void OpenSharedItemReturnsNullWhenSessionIsNull()
	{
		OutlookSession session = new(null);

		object? item = session.OpenSharedItem("nonexistent.eml");

		Assert.That(item, Is.Null);
	}

	/// <summary>
	/// Verifies that OutlookSession.GetStore returns null when the underlying
	/// session is null.
	/// </summary>
	/// <remarks>Constructs an OutlookSession with a null session, calls
	/// GetStore with a non-existent store name, and asserts the result is null.
	/// </remarks>
	[Test]
	public void GetStoreReturnsNullWhenSessionIsNull()
	{
		OutlookSession session = new(null);

		object? store = session.GetStore("doesnotexist.pst");

		Assert.That(store, Is.Null);
	}
}
