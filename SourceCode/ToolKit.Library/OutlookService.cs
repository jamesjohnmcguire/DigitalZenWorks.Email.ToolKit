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
		System.Reflection.MethodBase.GetCurrentMethod() !.DeclaringType);

	private IOutlookConnection? connection;
	private bool outlookStartedByThis;
	private IOutlookSession? session;

	/// <summary>
	/// Initializes a new instance of the <see cref="OutlookService"/> class.
	/// </summary>
	public OutlookService()
	{
	}

	/// <summary>
	/// Gets the Outlook session. This property will be null if the service
	/// is not connected to Outlook.
	/// </summary>
	public IOutlookSession? Session
	{
		get { return session; }
	}

	/// <summary>
	/// Gets a value indicating whether the service is currently connected to
	/// Outlook.
	/// </summary>
	public bool IsConnected
	{
		get { return connection != null && session != null; }
	}

	/// <summary>
	/// Checks if Outlook is installed on the system by looking for the
	/// installation path in the registry.
	/// </summary>
	/// <returns>True if Outlook is installed; otherwise, false.</returns>
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

	/// <summary>
	/// Connects to Outlook. If Outlook is not available, it will attempt
	/// to start a new instance.
	/// </summary>
	/// <param name="timeOutSeconds">The timeout in seconds.</param>
	/// <returns>True if the connection was successful; otherwise, false.
	/// </returns>
	public bool Connect(int timeOutSeconds = 10)
	{
		OutlookFactory factory = new();

		bool connected = Connect(factory, timeOutSeconds);

		return connected;
	}

	/// <summary>
	/// Disconnects from Outlook. If Outlook was started by the service,
	/// it will be quit.
	/// </summary>
	public void Disconnect()
	{
		try
		{
			if (connection != null && outlookStartedByThis == true)
			{
				connection.Quit();
			}
		}
		finally
		{
			connection = null;
			session = null;
			outlookStartedByThis = false;
		}
	}

	/// <summary>
	/// Checks if Outlook is currently running in the current user session.
	/// </summary>
	/// <returns>A boolean indicating whether Outlook is started.</returns>
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

	/// <summary>
	/// Connects to Outlook using the provided factory. If Outlook is not
	/// available, it will attempt to start a new instance.
	/// </summary>
	/// <param name="factory">The Outlook factory to use for creating
	/// connections.</param>
	/// <param name="timeOutSeconds">The timeout in seconds.</param>
	/// <returns>True if the connection was successful; otherwise, false.
	/// </returns>
	internal bool Connect(IOutlookFactory factory, int timeOutSeconds = 10)
	{
		bool connected = IsConnected;

		if (connected == false)
		{
			try
			{
				bool isAvailable = factory.CanCreateApplication(timeOutSeconds);

				if (isAvailable == true)
				{
					IOutlookConnection? candidateConnection =
						factory.CreateConnection();

					if (candidateConnection != null)
					{
						IOutlookSession? candidateSession =
							candidateConnection.Session;

						// Keep failed attempts from leaving a partial
						// connection that would prevent a later retry.
						if (candidateSession != null)
						{
							connection = candidateConnection;
							session = candidateSession;
							outlookStartedByThis = true;
							connected = true;
						}
					}
				}
			}
			catch (System.Exception exception) when
				(exception is COMException ||
				exception is InvalidOperationException)
			{
				Log.Error(exception);
			}
		}

		if (connected == false)
		{
			Log.Error("Outlook unavailable.");
		}

		return connected;
	}
}
