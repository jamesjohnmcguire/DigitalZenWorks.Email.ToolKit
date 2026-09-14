/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookProfile.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace ToolKit.Library;

using System.Collections.Generic;

/// <summary>
/// Represents an Outlook profile.
/// </summary>
public class OutlookProfile
{
	/// <summary>
	/// Gets the list of Outlook accounts in the profile.
	/// </summary>
	public IReadOnlyList<OutlookAccountInfo> Accounts { get; init; } = [];

	/// <summary>
	/// Gets the list of Outlook stores in the profile.
	/// </summary>
	public IReadOnlyList<OutlookStoreInfo> Stores { get; init; } = [];
}
