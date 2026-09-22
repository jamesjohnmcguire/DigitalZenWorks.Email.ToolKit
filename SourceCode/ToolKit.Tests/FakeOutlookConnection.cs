/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

using DigitalZenWorks.Email.ToolKit;

internal sealed class FakeOutlookConnection
    : IOutlookConnection
{
    public IOutlookSession Session { get; }

    public FakeOutlookConnection(
        IOutlookSession session)
    {
        Session = session;
    }

	/// <summary>
	/// Stub implementation of the Quit method. This should only be called by
	/// testing infrastructure and not by production code.
	/// </summary>
	public void Quit()
	{
	}
}
