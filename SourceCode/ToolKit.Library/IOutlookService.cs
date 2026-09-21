/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

/// <summary>
/// Interface for an Outlook service.
/// Public contract for connecting to and disconnecting from Outlook
/// and for obtaining a session abstraction.
/// </summary>
public interface IOutlookService
{
	/// <summary>
	/// Indicates whether there is an active connection/session.
	/// </summary>
	bool IsConnected { get; }

	/// <summary>
	/// Gets the active session abstraction or null when not connected.
	/// </summary>
	IOutlookSession? Session { get; }

	/// <summary>
	/// Connect to Outlook using the provided factory. This will attempt
	/// to attach to an existing Outlook instance or start a new one via
	/// the factory implementation.
	/// </summary>
	/// <param name="timeOutSeconds">Timeout for availability checks.</param>
	/// <returns>True when connected and a session is available.</returns>
	bool Connect(int timeOutSeconds = 10);

	/// <summary>
	/// Disconnect from Outlook. If Outlook was started by the service,
	/// the service may quit the application.
	/// </summary>
	void Disconnect();
}
