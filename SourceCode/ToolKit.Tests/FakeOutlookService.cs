/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

/// <summary>
/// Records CLI connection attempts and reports Outlook as unavailable.
/// </summary>
internal sealed class FakeOutlookService : IOutlookService
{
	/// <summary>
	/// Gets the number of connection attempts.
	/// </summary>
	public int ConnectCallCount { get; private set; }

	/// <inheritdoc/>
	public bool IsConnected => false;

	/// <summary>
	/// Gets the timeout supplied by the command dispatcher.
	/// </summary>
	public int? LastTimeoutSeconds { get; private set; }

	/// <inheritdoc/>
	public IOutlookSession? Session
	{
		get
		{
			SessionAccessCount++;
			return null;
		}
	}

	/// <summary>
	/// Gets the number of session requests after a failed connection.
	/// </summary>
	public int SessionAccessCount { get; private set; }

	/// <inheritdoc/>
	public bool Connect(int timeOutSeconds = 10)
	{
		ConnectCallCount++;
		LastTimeoutSeconds = timeOutSeconds;
		return false;
	}

	/// <inheritdoc/>
	public void Disconnect()
	{
	}
}
