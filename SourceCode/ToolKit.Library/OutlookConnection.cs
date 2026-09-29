/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Internal adapter that wraps an Outlook Application and exposes an
/// IOutlookSession. This class is intentionally internal; consumers
/// should use IOutlookService and IOutlookSession rather than this
/// concrete type.
/// </summary>
internal class OutlookConnection(Outlook.Application application)
		: IOutlookConnection
{
	private readonly Outlook.Application application = application;

	/// <summary>
	/// Gets the Outlook session.
	/// </summary>
	public IOutlookSession Session
	{
		get
		{
			return new OutlookSession(application.Session);
		}
	}

	/// <summary>
	/// Quits the Outlook application. This should only be called if the
	/// application was started by this instance.
	/// </summary>
	public void Quit()
	{
		if (application != null)
		{
			application.Quit();
		}
	}
}
