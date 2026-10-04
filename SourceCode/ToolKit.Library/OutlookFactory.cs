/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Runtime.InteropServices;
using global::Common.Logging;
using Microsoft.VisualBasic;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Factory class for creating connections to the Outlook application.
/// </summary>
public class OutlookFactory : IOutlookFactory
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod() !.DeclaringType);

	// This coordinates only temporary probes. No Outlook COM object is
	// cached here or shared across caller threads. A static coordinator is
	// needed because public Connect creates a new factory on every call.
	private static readonly OutlookApplicationProbe ApplicationProbe =
		new(ProbeApplication);

	private readonly OutlookApplicationProbe applicationProbe;

	/// <summary>
	/// Initializes a new instance of the <see cref="OutlookFactory"/> class.
	/// </summary>
	public OutlookFactory()
		: this(ApplicationProbe)
	{
	}

	/// <summary>
	/// Initializes a new instance of the <see cref="OutlookFactory"/> class
	/// with the supplied probe coordinator.
	/// </summary>
	/// <param name="applicationProbe">The activation probe coordinator.</param>
	internal OutlookFactory(OutlookApplicationProbe applicationProbe)
	{
		this.applicationProbe = applicationProbe;
	}

	/// <summary>
	/// Creates a connection to the Outlook application. If an existing
	/// instance of Outlook is running, it will connect to that instance;
	/// otherwise, it will attempt activation. The returned connection has no
	/// exclusive shutdown ownership, including when the probe started Outlook.
	/// Acquisition runs on the caller's thread and has no timeout.
	/// </summary>
	/// <returns>A connection to the Outlook application, or null if the
	/// connection could not be established.</returns>
	public IOutlookConnection? CreateConnection()
	{
		OutlookConnection? connection = null;

		Outlook.Application? application = CreateApplication();

		if (application != null)
		{
			connection = new(application);
		}

		return connection;
	}

	/// <summary>
	/// Waits for a temporary STA activation probe. Overlapping calls share
	/// the outstanding probe, including one that previously timed out.
	/// </summary>
	/// <param name="timeOutSeconds">The maximum requested wait in seconds
	/// for activation and cleanup. Must be between 0 and 2147483.</param>
	/// <returns>True if the probe completed successfully; false if it failed
	/// with a handled activation error or the caller stopped waiting.</returns>
	/// <remarks>
	/// Timeout does not cancel the worker. It may activate Outlook later.
	/// A completed probe is not cached as proof of future availability.
	/// Application and session acquisition by CreateConnection are separate,
	/// run on the caller's thread, and are not covered by this timeout.
	/// The probe's COM reference must never escape its temporary apartment.
	/// </remarks>
	public bool CanCreateApplication(int timeOutSeconds)
	{
		bool isAvailable =
			applicationProbe.CanCreateApplication(timeOutSeconds);

		return isAvailable;
	}

	private static void ProbeApplication()
	{
		Outlook.Application? tryApplication = null;

		try
		{
			tryApplication = new Outlook.Application();
		}
		finally
		{
			if (tryApplication != null)
			{
				// Balance this acquisition on its STA. Do not force all
				// references to a potentially shared RCW to be released.
				Marshal.ReleaseComObject(tryApplication);
			}
		}
	}

	private static Outlook.Application? ConnectToExistingOutlook()
	{
		Outlook.Application? application = null;

		try
		{
#if NET5_0_OR_GREATER
			application = Interaction.GetObject(null, "Outlook.Application")
					as Outlook.Application;
#elif !NETSTANDARD2_0_OR_GREATER
			application = Marshal.GetActiveObject("Outlook.Application")
					as Outlook.Application;
#endif
		}
		catch (Exception exception) when
			(exception is COMException ||
			exception is InvalidOperationException)
		{
			Log.Debug(
				"Could not attach to an existing Outlook instance.");
			Log.Debug(exception.ToString());
		}

		return application;
	}

	private static Outlook.Application? CreateApplication()
	{
		Outlook.Application? application = null;

		try
		{
			application = ConnectToExistingOutlook();

			if (application == null)
			{
				application = new Outlook.Application();
			}
		}
		catch (Exception exception) when
			(exception is COMException ||
			exception is InvalidOperationException)
		{
			Log.Error(exception.ToString());
		}

		return application;
	}
}
