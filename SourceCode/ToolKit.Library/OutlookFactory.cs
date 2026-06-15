/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using global::Common.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

#nullable enable

public static class OutlookFactory
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

	public static bool IsOutlookAvailable(int timeOutSeconds)
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
			catch (Exception ex)
			{
				exception = ex;
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
}
