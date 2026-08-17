/////////////////////////////////////////////////////////////////////////////
// <copyright file="IOutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

public interface IOutlookConnection
{
	IOutlookSession Session { get; }
}
