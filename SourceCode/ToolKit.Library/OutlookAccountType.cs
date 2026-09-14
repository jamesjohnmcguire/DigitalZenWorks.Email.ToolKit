/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookAccountType.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace ToolKit.Library;

/// <summary>
/// Represents the type of an Outlook account.
/// </summary>
public enum OutlookAccountType
{
	/// <summary>
	/// The Outlook account type is unknown.
	/// </summary>
	Unknown,

	/// <summary>
	/// The Outlook account type is Exchange.
	/// </summary>
	Exchange,

	/// <summary>
	/// The Outlook account type is HTTP.
	/// </summary>
	Http,

	/// <summary>
	/// The Outlook account type is IMAP.
	/// </summary>
	Imap,

	/// <summary>
	/// The Outlook account type is POP3.
	/// </summary>
	Pop3,

	/// <summary>
	/// The Outlook account type is other.
	/// </summary>
	Other
}
