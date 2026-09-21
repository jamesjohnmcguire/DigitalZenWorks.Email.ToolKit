/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

#nullable enable

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using global::Common.Logging;
using Microsoft.Office.Interop.Outlook;
#if NETFRAMEWORK || NETSTANDARD2_0_OR_GREATER || NET6_0_OR_GREATER
using Microsoft.Win32;
#endif
using Microsoft.VisualBasic;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Public service that manages Outlook startup/attach lifecycle and
/// exposes an IOutlookSession for performing Outlook operations.
/// New code should prefer this surface for session acquisition and
/// lifecycle control. The legacy OutlookAccount singleton remains for
/// compatibility but new code should avoid using it directly.
/// </summary>
public class OutlookService : IOutlookService
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

	private Application? application;
	private IOutlookConnection? connection;
	private bool outlookStartedByThis;
	private bool attachedToExistingOutlook;
	private IOutlookSession? session;

	public OutlookService()
	{
	}

	public IOutlookSession? Session
	{
		get { return session; }
	}

	public bool IsConnected
	{
		get { return connection != null; }
	}

	public bool Connect(IOutlookFactory factory, int timeOutSeconds = 10)
	{
		bool connected = false;

		if (connection == null)
		{
			application = ConnectToExistingOutlook();

			if (application != null)
			{
				connection = new OutlookConnection(application);
				attachedToExistingOutlook = true;
			}
			else
			{
				try
				{
					bool isAvailable =
						factory.IsOutlookAvailable(timeOutSeconds);

					if (isAvailable == true)
					{
						connection = factory.CreateConnection();

						if (connection != null)
						{
							session = connection.Session;
							outlookStartedByThis = true;
						}
					}
				}
				catch (System.Exception exception) when
					(exception is COMException ||
					exception is InvalidOperationException)
				{
					Log.Error(exception);
					connection = null;
				}
			}
		}

		if (connection != null)
		{
			session = connection.Session;

			connected = true;
		}
		else
		{
			Log.Error("Outlook unavailable.");
		}

		return connected;
	}

	public void Disconnect()
	{
		if (outlookStartedByThis == true && application != null)
		{
			application.Quit();
		}
	}

	public static bool IsOutlookInstalled()
	{
		bool installed = false;

#if NETFRAMEWORK || NETSTANDARD2_0_OR_GREATER || NET6_0_OR_GREATER
		string registryPath =
			@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE";
		using RegistryKey? key =
			Registry.LocalMachine.OpenSubKey(registryPath);

		if (key != null)
		{
			installed = true;
		}
		else
		{
			// 32-bit Outlook on 64-bit Windows
			registryPath = @"SOFTWARE\WOW6432Node\Microsoft\Windows\" +
				@"CurrentVersion\App Paths\OUTLOOK.EXE";
			using RegistryKey? wowKey =
				Registry.LocalMachine.OpenSubKey(registryPath);

			if (wowKey != null)
			{
				installed = true;
			}
		}
#endif

		return installed;
	}

	private bool IsOutlookStarted()
	{
		bool started = false;

		if (application == null)
		{
			Process[] existing = Process.GetProcessesByName("OUTLOOK");
			int count = existing.Length;

			if (count == 0)
			{
				started = true;
			}
		}

		return started;
	}

	private static Application? ConnectToExistingOutlook()
	{
		Application? application = null;

		try
		{
#if NET5_0_OR_GREATER
			application = Interaction.GetObject(null, "Outlook.Application")
					as Application;
#elif !NETSTANDARD2_0_OR_GREATER
			application = Marshal.GetActiveObject("Outlook.Application")
					as Application;
#endif
		}
		catch (COMException exception)
		{
			Log.Debug(
				"Could not attach to an existing Outlook instance.");
			Log.Debug(exception.ToString());
		}

		return application;
	}
}
