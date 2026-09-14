/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookStoreInfo.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace ToolKit.Library;

/// <summary>
/// Represents information about an Outlook store.
/// </summary>
public class OutlookStoreInfo
{
	/// <summary>
	/// Gets the display name of the Outlook store.
	/// </summary>
	public string? DisplayName { get; init; }

	/// <summary>
	/// Gets the file path of the Outlook store.
	/// </summary>
	public string? FilePath { get; init; }

	/// <summary>
	/// Gets the unique identifier of the Outlook store.
	/// </summary>
	public string? StoreId { get; init; }

	/// <summary>
	/// Gets a value indicating whether the Outlook store is a data
	/// file store.
	/// </summary>
	public bool IsDataFileStore { get; init; }

	/// <summary>
	/// Gets a value indicating whether the Outlook store is currently
	/// open.
	/// </summary>
	public bool IsOpen { get; init; }
}
