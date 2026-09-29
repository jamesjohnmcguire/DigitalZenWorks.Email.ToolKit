/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Runtime.InteropServices;
using System.Threading;
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

	/// <summary>
	/// Creates a connection to the Outlook application. If an existing
	/// instance of Outlook is running, it will connect to that instance;
	/// otherwise, it will create a new instance.
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
	/// Checks if Outlook is available by attempting to create an instance of
	/// the Outlook application within a specified timeout period.
	/// </summary>
	/// <param name="timeOutSeconds">The timeout in seconds.</param>
	/// <returns>True if Outlook is available; otherwise, false.</returns>
	public bool IsOutlookAvailable(int timeOutSeconds)
	{
		bool isAvailable = false;
		Outlook.Application? tryApplication = null;

		Exception? exception = null;

		TimeSpan timeOutSpan = TimeSpan.FromSeconds(timeOutSeconds);

		using ManualResetEvent completed = new(initialState: false);

		void CreateOutlookApplication()
		{
			try
			{
				tryApplication = new Outlook.Application();
			}
			catch (System.Exception innerException) when
				(exception is COMException ||
				exception is InvalidOperationException)
			{
				exception = innerException;
				Log.Error(exception.ToString());
			}
			finally
			{
				if (tryApplication != null)
				{
					Marshal.FinalReleaseComObject(tryApplication);
				}

				completed.Set();
			}
		}

		Thread staThread = new(CreateOutlookApplication);

		staThread.SetApartmentState(ApartmentState.STA);
		staThread.IsBackground = true;
		staThread.Start();

		bool finished = completed.WaitOne(timeOutSpan);

		if (finished == true && exception == null)
		{
			isAvailable = true;
		}

		return isAvailable;
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
