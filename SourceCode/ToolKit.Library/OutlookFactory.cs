/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

#nullable enable

public static class OutlookFactory
{
	public static Outlook.Application? TryCreateOutlookApplication(
		TimeSpan timeout)
	{
		Outlook.Application? application = null;
		Outlook.Application? tryApplication = null;

		Exception? exception = null;

		using ManualResetEvent completed = new(initialState: false);

		Thread staThread = new(() =>
		{
			try
			{
				tryApplication = new Outlook.Application();
			}
			catch (Exception ex)
			{
				exception = ex;
			}
			finally

			{
				completed.Set();
			}
		});

		staThread.SetApartmentState(ApartmentState.STA);
		staThread.IsBackground = true;
		staThread.Start();

		bool finished = completed.WaitOne(timeout);

		if (finished == true && exception == null)
		{
			application = tryApplication;
		}

		return application;
	}
}
