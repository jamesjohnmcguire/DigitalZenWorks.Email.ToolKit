/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookAccountSessionAdapter.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using Microsoft.Office.Interop.Outlook;

/// <summary>
/// Internal adapter that implements IOutlookSession by delegating to
/// the legacy OutlookAccount singleton. This allows code written to
/// consume IOutlookSession to interoperate with existing OutlookAccount
/// implementations without changing the legacy API.
/// </summary>
internal sealed class OutlookAccountSessionAdapter : IOutlookSession
{
	/// <summary>
	/// Opens a shared item by delegating to the OutlookAccount session.
	/// </summary>
	/// <param name="filePath">The path to the shared item file.</param>
	/// <returns>The opened item or null.</returns>
	public object? OpenSharedItem(string filePath)
	{
		NameSpace? session = OutlookAccount.Instance.Session;

		if (session == null)
		{
			return null;
		}

		object? item = session.OpenSharedItem(filePath);
		return item;
	}

	/// <summary>
	/// Removes a store by delegating to OutlookAccount.RemoveStore.
	/// </summary>
	/// <param name="path">The path to the store to remove.</param>
	/// <returns>True if removed, false otherwise.</returns>
	public bool RemoveStore(string path)
	{
		bool result = OutlookAccount.Instance.RemoveStore(path);
		return result;
	}

	/// <summary>
	/// Gets an item from its EntryID by delegating to the OutlookAccount
	/// session.
	/// </summary>
	/// <param name="entryId">The EntryID of the item to retrieve.</param>
	/// <returns>The retrieved item or null.</returns>
	public object? GetItemFromId(string entryId)
	{
		object? item = null;
		NameSpace? session = OutlookAccount.Instance.Session;

		if (session != null)
		{
			item = session.GetItemFromID(entryId);
		}

		return item;
	}
}
