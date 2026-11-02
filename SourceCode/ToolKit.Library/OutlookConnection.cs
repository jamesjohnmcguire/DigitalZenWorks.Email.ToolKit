/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Internal adapter that wraps an Outlook Application and exposes an
/// IOutlookSession. This class is intentionally internal; consumers
/// should use IOutlookService and IOutlookSession rather than this
/// concrete type.
/// </summary>
internal sealed class OutlookConnection
		: IOutlookConnection, IDisposable
{
	private Outlook.Application? application;
	private IOutlookSession? session;

	/// <summary>
	/// Initializes a new instance of the <see cref="OutlookConnection"/> class.
	/// </summary>
	/// <param name="application">The Outlook application to wrap.</param>
	public OutlookConnection(Outlook.Application application)
	{
		this.application = application;
		session = new OutlookSession(application.Session);
	}

	/// <summary>
	/// Initializes a new instance of the <see cref="OutlookConnection"/> class.
	/// </summary>
	/// <param name="session">The session to expose.</param>
	/// <remarks>This constructor is intended for testing purposes only.
	/// </remarks>
	internal OutlookConnection(IOutlookSession session)
	{
		this.application = null;
		this.session = session;
	}

	/// <summary>
	/// Gets the stable session, or null after this connection is disposed.
	/// </summary>
	public IOutlookSession? Session
	{
		get { return session; }
	}

	/// <summary>
	/// Relinquishes this connection's managed references without quitting
	/// Outlook or invalidating COM wrappers held by another client.
	/// </summary>
	/// <remarks>
	/// Activation, including the temporary probe, cannot establish exclusive
	/// ownership of Outlook. Treat every connection as shared or unknown.
	/// Sessions and COM objects can escape through existing public APIs, so
	/// forcing their RCW counts to zero here could invalidate other callers.
	/// The runtime releases those wrappers when no managed owners remain.
	/// This is not deterministic release of every underlying COM reference.
	/// </remarks>
	public void Dispose()
	{
		session = null;
		application = null;
		GC.SuppressFinalize(this);
	}

	/// <summary>
	/// Explicitly requests Outlook shutdown. The caller must have authority
	/// to close the application. Disposal never calls this method.
	/// </summary>
	public void Quit()
	{
		if (application != null)
		{
			application.Quit();
		}
	}
}
