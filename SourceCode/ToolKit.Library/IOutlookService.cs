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
	/// Gets a value indicating whether there is an active connection/session.
	/// </summary>
	bool IsConnected { get; }

	/// <summary>
	/// Gets the active session abstraction or null when not connected.
	/// </summary>
	IOutlookSession? Session { get; }

	/// <summary>
	/// Connects to Outlook by attaching to an existing instance or attempting
	/// activation. Reuses the connection when already connected.
	/// </summary>
	/// <param name="timeOutSeconds">Maximum requested wait, in seconds, for the
	/// activation probe to finish,including its cleanup.</param>
	/// <returns>True when connected and a session is available.</returns>
	/// <remarks>
	/// A timeout stops waiting; the probe may still complete later.
	/// Subsequent application and session acquisition run on the caller's
	/// thread and have no timeout. This is not a total Connect deadline.
	/// Reuse this service on the same STA thread for the workflow's lifetime.
	/// Concurrent Connect and Disconnect calls are not supported.
	/// </remarks>
	bool Connect(int timeOutSeconds = 10);

	/// <summary>
	/// Clears this service's connection and session without quitting Outlook.
	/// The connection is disposed if it supports IDisposable.
	/// </summary>
	void Disconnect();
}
