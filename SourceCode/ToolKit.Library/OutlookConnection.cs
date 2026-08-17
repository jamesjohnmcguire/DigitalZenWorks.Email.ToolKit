/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

public class OutlookConnection
	: IOutlookConnection
{
	private readonly Outlook.Application application;

	public OutlookConnection(Outlook.Application application)
	{
		this.application = application;
	}

	public IOutlookSession Session
	{
		get
		{
			return new OutlookSession(application.Session);
		}
	}
}
