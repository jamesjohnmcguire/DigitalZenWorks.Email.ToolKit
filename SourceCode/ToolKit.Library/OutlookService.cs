/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
#if NETFRAMEWORK || NETSTANDARD2_0_OR_GREATER || NET6_0_OR_GREATER
using Microsoft.Win32;
#endif
using Microsoft.VisualBasic;

public class OutlookService : IOutlookService
{
	private static readonly OutlookService InternalInstance = new();

	private Application? application;
	private bool outlookStartedByThis;
	private bool attachedToExistingOutlook;
	private OutlookSession? session;

	private OutlookService()
	{
	}

    public static OutlookService Instance
    {
        get { return InternalInstance; }
    }

	public OutlookSession? Session
	{
		get { return session; }
	}

	public bool Connect(int timeOutSeconds = 10)
	{
		bool connected = false;

		if (application == null)
		{
			application = ConnectToExistingOutlook();

			if (application != null)
			{
				attachedToExistingOutlook = true;
			}
			else
			{
				application =
					OutlookFactory.TryCreateOutlookApplication(timeOutSeconds);

				if (application != null)
				{
					outlookStartedByThis = true;
				}
			}
		}

		if (application != null)
		{
			session = new OutlookSession(application);

			connected = true;
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
			application = Interaction.GetObject(null, "Outlook.Application")
					as Application;
		}
		catch
		{
		}

		return application;
	}
}
