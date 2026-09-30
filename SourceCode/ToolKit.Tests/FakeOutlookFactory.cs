/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

/// <summary>
/// A fake implementation of the IOutlookFactory interface for testing purposes.
/// </summary>
internal sealed class FakeOutlookFactory : IOutlookFactory
{
	/// <summary>
	/// Gets or sets a value indicating whether Outlook is available.
	/// </summary>
	public bool IsAvailable { get; set; }

	/// <summary>
	/// Gets or sets the Outlook connection to be returned by the
	/// CreateConnection method.
	/// </summary>
	public IOutlookConnection? Connection { get; set; }

	/// <summary>
	/// Gets the number of times the CreateConnection method has been called.
	/// </summary>
	public int CreateConnectionCallCount { get; private set; }

	/// <summary>
	/// Gets the number of times the IsOutlookAvailable method has been called.
	/// </summary>
	public int IsOutlookAvailableCallCount { get; private set; }

	/// <summary>
	/// Creates a connection to Outlook.
	/// </summary>
	/// <returns>The Outlook connection.</returns>
	public IOutlookConnection? CreateConnection()
	{
		CreateConnectionCallCount++;

		return Connection;
	}

	/// <summary>
	/// Determines if Outlook is available.
	/// </summary>
	/// <param name="timeoutSeconds">The timeout in seconds.</param>
	/// <returns>A boolean value indicating whether Outlook is available.
	/// </returns>
	public bool CanCreateApplication(int timeoutSeconds)
	{
		IsOutlookAvailableCallCount++;

		return IsAvailable;
	}
}
