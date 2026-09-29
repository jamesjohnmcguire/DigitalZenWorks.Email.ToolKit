/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using global::Common.Logging;
#if NETFRAMEWORK || NETSTANDARD2_0_OR_GREATER || NET6_0_OR_GREATER
using Microsoft.Win32;
#endif

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

	private IOutlookConnection? connection;
	private bool outlookStartedByThis;
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

	public bool Connect(int timeOutSeconds = 10)
	{
		OutlookFactory factory = new();

		bool connected = Connect(factory, timeOutSeconds);

		return connected;
	}

	internal bool Connect(IOutlookFactory factory, int timeOutSeconds = 10)
	{
		bool connected = false;

		if (connection != null)
		{
			connected = true;
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

		if (connection != null)
		{
			connected = true;
		}
		else
		{
			Log.Error("Outlook unavailable.");
		}

		return connected;
	}

	/// <summary>
	/// Disconnects from Outlook. If Outlook was started by the service,
	/// it will be quit.
	/// </summary>
	public void Disconnect()
	{
		if (connection != null && outlookStartedByThis == true)
		{
			connection.Quit();
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

	internal static bool IsOutlookStarted()
	{
		bool started = false;

		Process[] existing = Process.GetProcessesByName("OUTLOOK");
		int count = existing.Length;

		if (count > 0)
		{
			started = true;
		}

		return started;
	}
}
