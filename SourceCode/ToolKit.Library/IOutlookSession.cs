/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookSession.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

#nullable enable

/// <summary>
/// Interface for an Outlook session.
/// </summary>
public interface IOutlookSession
{
#if FUTURE
	IEnumerable<IOutlookStore> Stores { get; }

	public bool AddStore(string path);

	IOutlookStore GetStore(string path);
#endif

	/// <summary>
	/// Opens a shared item in Outlook.
	/// </summary>
	/// <param name="filePath">The path to the file to open.</param>
	/// <returns>The opened item, or null if the item could not be opened.
	/// </returns>
	public object? OpenSharedItem(string filePath);

	/// <summary>
	/// Removes a store from the Outlook session.
	/// </summary>
	/// <param name="path">The path to the store to remove.</param>
	/// <returns>true if the store was removed, false otherwise.</returns>
	public bool RemoveStore(string path);
}
