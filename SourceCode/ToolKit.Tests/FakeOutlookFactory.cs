/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

using Outlook = Microsoft.Office.Interop.Outlook;

#nullable enable

public sealed class FakeOutlookFactory : IOutlookFactory
{
	public bool IsAvailable { get; set; }

	public IOutlookConnection? Connection { get; set; }

	public int CreateConnectionCallCount { get; private set; }

	public int IsOutlookAvailableCallCount { get; private set; }

	public IOutlookConnection? CreateConnection()
	{
		CreateConnectionCallCount++;

		return Connection;
	}

	public bool IsOutlookAvailable(int timeoutSeconds)
	{
		IsOutlookAvailableCallCount++;

		return IsAvailable;
	}
}
