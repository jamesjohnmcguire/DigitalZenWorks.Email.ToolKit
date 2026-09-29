/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

using DigitalZenWorks.Email.ToolKit;

/// <summary>
/// A fake implementation of the IOutlookConnection interface for testing
/// purposes.
/// </summary>
internal sealed class FakeOutlookConnection
	: IOutlookConnection
{
	/// <summary>
	/// Initializes a new instance of the <see cref="FakeOutlookConnection"/>
	/// class.
	/// </summary>
	public FakeOutlookConnection()
	{
	}

	/// <summary>
	/// Initializes a new instance of the <see cref="FakeOutlookConnection"/>
	/// class.
	/// </summary>
	/// <param name="session">The Outlook session.</param>
	public FakeOutlookConnection(IOutlookSession session)
	{
		Session = session;
	}

	/// <summary>
	/// Gets the Outlook session.
	/// </summary>
	public IOutlookSession Session { get; }

	/// <summary>
	/// Stub implementation of the Quit method. This should only be called by
	/// testing infrastructure and not by production code.
	/// </summary>
	public void Quit()
	{
	}
}
