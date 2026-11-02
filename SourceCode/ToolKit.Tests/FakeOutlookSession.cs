/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookSession.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using DigitalZenWorks.Email.ToolKit;
using Microsoft.Office.Interop.Outlook;

/// <summary>
/// A fake implementation of the IOutlookSession interface for testing purposes.
/// </summary>
internal sealed class FakeOutlookSession
	: IOutlookSession
{
	/// <summary>
	/// Gets a store from the Outlook session by its path.
	/// </summary>
	/// <param name="path">The path of the store to get.</param>
	/// <returns>The retrieved store.</returns>
	public static Store? GetStore(string path)
	{
		return null;
	}

	/// <summary>
	/// Gets an item from the Outlook session by its entry ID.
	/// </summary>
	/// <param name="entryId">The entry ID of the item to get.</param>
	/// <returns>The retrieved item.</returns>
	public object? GetItemFromId(string entryId)
	{
		return new object();
	}

	/// <summary>
	/// Opens a shared item in Outlook.
	/// </summary>
	/// <param name="filePath">The path of the file to open.</param>
	/// <returns>The opened item.</returns>
	public object? OpenSharedItem(string filePath)
	{
		return new object();
	}

	/// <summary>
	/// Removes a store from the Outlook session.
	/// </summary>
	/// <param name="path">The path of the store to remove.</param>
	/// <returns>A boolean value indicating whether the store was removed.
	/// </returns>
	public bool RemoveStore(string path)
	{
		return true;
	}
}
