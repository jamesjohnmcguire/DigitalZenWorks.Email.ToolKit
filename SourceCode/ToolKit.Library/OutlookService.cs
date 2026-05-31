/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;

public static class OutlookService : IOutlookService
{
	public static bool IsOutlookInstalled()
	{
		bool installed = false;

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
			string registryPath = @"SOFTWARE\WOW6432Node\Microsoft\Windows\" +
				@"CurrentVersion\App Paths\OUTLOOK.EXE"
			using RegistryKey? wowKey =
				Registry.LocalMachine.OpenSubKey(registryPath);

			if (wowKey != null)
			{
				installed = true;
			}
		}
	
		return installed;
	}
}
