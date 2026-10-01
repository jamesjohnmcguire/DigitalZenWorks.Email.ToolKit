/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookSessionTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Tests session behavior without an underlying COM namespace.
/// </summary>
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
	/// Verifies either creation option returns null without an Outlook session.
	/// </summary>
	/// <param name="createIfMissing">Whether creation is allowed.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void GetStoreReturnsNullWhenSessionIsNull(bool createIfMissing)
	{
		OutlookSession session = new(null);

		object? store = session.GetStore("doesnotexist.pst", createIfMissing);

		Assert.That(store, Is.Null);
	}

	/// <summary>
	/// Verifies folder wrappers are absent without a namespace.
	/// </summary>
	[Test]
	public void GetFolderFromIdReturnsNullWhenSessionIsNull()
	{
		OutlookSession session = new(null);

		Assert.That(session.GetFolderFromId("entry", "store"), Is.Null);
	}

	/// <summary>
	/// Verifies raw folder lookup is absent without a namespace.
	/// </summary>
	[Test]
	public void GetFolderFromIdInternalReturnsNullWhenSessionIsNull()
	{
		OutlookSession session = new(null);

		Assert.That(
			session.GetFolderFromIdInternal("entry", "store"), Is.Null);
	}

	/// <summary>
	/// Verifies mail wrapping is absent without a namespace.
	/// </summary>
	[Test]
	public void OpenMailItemFileReturnsNullWhenSessionIsNull()
	{
		OutlookSession session = new(null);

		Assert.That(session.OpenMailItemFile("missing.msg"), Is.Null);
	}

	/// <summary>
	/// Verifies store removal fails without a namespace.
	/// </summary>
	/// <param name="path">A path accepted by the path normalizer.</param>
	[TestCase("missing.pst")]
	[TestCase("missing.PST")]
	[TestCase("missing")]
	public void RemoveStoreReturnsFalseWhenSessionIsNull(string path)
	{
		OutlookSession session = new(null);

		Assert.That(session.RemoveStore(path), Is.False);
	}

	/// <summary>
	/// Verifies a missing store object cannot be removed.
	/// </summary>
	[Test]
	public void RemoveStoreReturnsFalseWhenStoreIsNull()
	{
		OutlookSession session = new(null);

		Assert.That(session.RemoveStore((Outlook.Store)null!), Is.False);
	}
}
