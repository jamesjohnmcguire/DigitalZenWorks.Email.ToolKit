/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

#nullable enable

using Outlook = Microsoft.Office.Interop.Outlook;

public interface IOutlookFactory
{
	public Outlook.Application? CreateApplication();

	public bool IsOutlookAvailable(int timeOutSeconds);
}
