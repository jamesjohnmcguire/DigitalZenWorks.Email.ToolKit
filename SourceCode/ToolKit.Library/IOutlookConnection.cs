/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

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
	/// Explicitly requests Outlook shutdown. The caller must have authority
	/// to close the application, which may also be in use by other clients.
	/// Normal service disconnection does not call this method.
	/// </summary>
	void Quit();
}
