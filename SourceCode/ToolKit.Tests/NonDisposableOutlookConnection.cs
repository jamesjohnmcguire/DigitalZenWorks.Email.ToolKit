/////////////////////////////////////////////////////////////////////////////
// <copyright file="NonDisposableOutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

/// <summary>
/// Represents the original public connection contract without IDisposable.
/// </summary>
internal sealed class NonDisposableOutlookConnection : IOutlookConnection
{
	/// <summary>
	/// Gets the session returned to the service.
	/// </summary>
	public IOutlookSession Session { get; } = new FakeOutlookSession();

	/// <summary>
	/// Gets the number of application shutdown requests.
	/// </summary>
	public int QuitCallCount { get; private set; }

	/// <summary>
	/// Records application shutdown requests.
	/// </summary>
	public void Quit()
	{
		QuitCallCount++;
	}
}
