/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookProfileService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace ToolKit.Library;

/// <summary>
/// Defines a service for retrieving Outlook profile information.
/// </summary>
public interface IOutlookProfileService
{
	/// <summary>
	/// Retrieves information about the current Outlook profile.
	/// </summary>
	/// <returns>An <see cref="OutlookProfile"/> object containing information
	/// about the current Outlook profile.</returns>
	OutlookProfile GetProfileInfo();
}
