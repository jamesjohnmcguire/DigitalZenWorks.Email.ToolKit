/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

/// <summary>
/// Interface for an Outlook factory.
/// </summary>
internal interface IOutlookFactory
{
	/// <summary>
	/// Creates a connection to Outlook.
	/// </summary>
	/// <returns>Reference to the Outlook connection, or null if the connection
	/// could not be established.</returns>
	public IOutlookConnection? CreateConnection();

	/// <summary>
	/// Determines if Outlook is available.
	/// </summary>
	/// <param name="timeOutSeconds">The requested wait for the activation
	/// probe, including cleanup, in seconds. Timeout does not cancel the
	/// probe or bound subsequent connection acquisition.</param>
	/// <returns>true if Outlook is available, false otherwise.</returns>
	public bool CanCreateApplication(int timeOutSeconds);
}
