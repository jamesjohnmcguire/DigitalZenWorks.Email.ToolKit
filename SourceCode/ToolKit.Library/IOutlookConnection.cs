/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

#nullable enable

/// <summary>
/// Interface for an Outlook connection.
/// </summary>
public interface IOutlookConnection
{
	/// <summary>
	/// Gets the Outlook session.
	/// </summary>
	IOutlookSession? Session { get; }

	/// <summary>
	/// Quits the Outlook application. This should only be called if the
	/// application was started by this instance.
	/// </summary>
	void Quit();
}
