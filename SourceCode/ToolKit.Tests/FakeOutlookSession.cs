/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookSession.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

using DigitalZenWorks.Email.ToolKit;

internal sealed class FakeOutlookSession
	: IOutlookSession
{
	public object? OpenSharedItem(string filePath)
	{
		return new object();
	}

	public bool RemoveStore(string path)
	{
		return true;
	}
}
