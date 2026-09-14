/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookAccountInfo.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace ToolKit.Library;

/// <summary>
/// Represents information about an Outlook account.
/// </summary>
public class OutlookAccountInfo
{
	/// <summary>
	/// Gets the display name of the Outlook account.
	/// </summary>
	public string? DisplayName { get; init; }

	/// <summary>
	/// Gets the user name associated with the Outlook account.
	/// </summary>
	public string? UserName { get; init; }

	/// <summary>
	/// Gets the SMTP address associated with the Outlook account.
	/// </summary>
	public string? SmtpAddress { get; init; }

	/// <summary>
	/// Gets the type of the Outlook account.
	/// </summary>
	public OutlookAccountType AccountType { get; init; }

	/// <summary>
	/// Gets the name of the delivery store associated with the Outlook account.
	/// </summary>
	public string? DeliveryStoreName { get; init; }

	/// <summary>
	/// Gets the path of the delivery store associated with the Outlook account.
	/// </summary>
	public string? DeliveryStorePath { get; init; }
}
