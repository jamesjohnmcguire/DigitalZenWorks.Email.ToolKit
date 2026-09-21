/////////////////////////////////////////////////////////////////////////////
// <copyright file="ThrowingFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;

/// <summary>
/// A factory that simulates Outlook being available but CreateConnection
/// throwing an exception.
/// </summary>
internal class ThrowingFactory : IOutlookFactory
{
	/// <summary>
	/// Simulates Outlook being available but CreateConnection throwing an
	/// exception.
	/// </summary>
	/// <returns>The created Outlook connection.</returns>
	/// <exception cref="InvalidOperationException">An error occurred while
	/// creating the Outlook connection.</exception>
	public IOutlookConnection? CreateConnection()
	{
		throw new InvalidOperationException("CreateConnection failed");
	}

	/// <summary>
	/// Simulates Outlook being available but CreateConnection failing.
	/// </summary>
	/// <param name="timeOutSeconds">The time out period in seconds.</param>
	/// <returns>true if Outlook is available; otherwise, false.</returns>
	public bool IsOutlookAvailable(int timeOutSeconds)
	{
		return true;
	}
}
