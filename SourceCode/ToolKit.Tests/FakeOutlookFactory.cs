/////////////////////////////////////////////////////////////////////////////
// <copyright file="EmailToolKitTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

public sealed class FakeOutlookFactory : IOutlookFactory
{
	public bool IsAvailable { get; set; }

	public int CallCount { get; private set; }

	public Application? CreateApplication()
	{
		CreateCallCount++;

		Application? fakeApplication = null;

		return fakeApplication;
	}

	public bool IsOutlookAvailable(int timeoutSeconds)
	{
		CallCount++;

		return IsAvailable;
	}
}
