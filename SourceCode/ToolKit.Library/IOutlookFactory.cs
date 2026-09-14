/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

#nullable enable

/// <summary>
/// Interface for an Outlook factory.
/// </summary>
public interface IOutlookFactory
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
	/// <param name="timeOutSeconds">The time out period in seconds.</param>
	/// <returns>true if Outlook is available, false otherwise.</returns>
	public bool IsOutlookAvailable(int timeOutSeconds);
}
