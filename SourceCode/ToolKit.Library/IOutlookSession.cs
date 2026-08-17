/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookSession.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

public interface IOutlookSession
{
#if FUTURE
	IEnumerable<IOutlookStore> Stores { get; }

	public bool AddStore(string path);

	IOutlookStore GetStore(string path);
#endif

	public object? OpenSharedItem(string filePath);

	public bool RemoveStore(string path);
}
