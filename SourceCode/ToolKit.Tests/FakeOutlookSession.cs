/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookSession.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using DigitalZenWorks.Email.ToolKit;
using Microsoft.Office.Interop.Outlook;

internal sealed class FakeOutlookSession
	: IOutlookSession
{

	public Store? GetStore(string path)
	{
		return null;
	}

	public object? GetItemFromId(string entryId)
	{
		return new object();
	}

	public object? OpenSharedItem(string filePath)
	{
		return new object();
	}

	public bool RemoveStore(string path)
	{
		return true;
	}

	public MAPIFolder? GetFolderFromIdInternal(string entryId, string storeId)
	{
		return null;
	}
}
