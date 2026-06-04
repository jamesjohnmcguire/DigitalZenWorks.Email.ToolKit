/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookSession.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.IO;
using global::Common.Logging;
using Microsoft.Office.Interop.Outlook;

public class OutlookSession
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

	private readonly NameSpace? session;

	public OutlookSession(Application application)
	{
		session = application.Session;
	}

	/// <summary>
	/// Create a new pst storage file.
	/// </summary>
	/// <param name="path">The path to the pst file.</param>
	/// <returns>A store object.</returns>
	public Store GetStore(string path)
	{
		Store newPst = null;

		path = Path.GetFullPath(path);

		string extension = Path.GetExtension(path);

		if (!extension.Equals(".pst", StringComparison.OrdinalIgnoreCase))
		{
			// Attempt to fix mistaken or missing file extension.
			path += ".pst";
		}

		// If the .pst file does not exist, Microsoft Outlook creates it.
		session.AddStore(path);

		int total = session.Stores.Count;

		for (int index = 1; index <= total; index++)
		{
			Store store = null;

			try
			{
				store = session.Stores[index];
			}
			catch (UnauthorizedAccessException exception)
			{
				Log.Error(exception.ToString());
			}

			if (store == null)
			{
				Log.Warn("Enumerating stores - store is null");
			}
			else
			{
				string filePath = store.FilePath;

				if (!string.IsNullOrWhiteSpace(filePath) &&
					filePath.Equals(
						path, StringComparison.OrdinalIgnoreCase))
				{
					newPst = store;
					break;
				}
			}
		}

		if (newPst == null)
		{
			Log.Warn("Store not found: " + path);
		}

		return newPst;
	}

	/// <summary>
	/// Removes a store from Outlook.
	/// </summary>
	/// <param name="path">The path to the pst file.</param>
	/// <returns>remove result.</returns>
	public bool RemoveStore(string path)
	{
		bool result = false;

		Log.Info("Begin to Removing store: " + path);

		path = Path.GetFullPath(path);
		string extension = Path.GetExtension(path);

		if (!extension.Equals(".pst", StringComparison.OrdinalIgnoreCase))
		{
			// Attempt to fix mistaken or missing file extension.
			path += ".pst";
		}

		Store store = GetStore(path);

		if (store != null)
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			session.RemoveStore(rootFolder);

			Log.Info("Store removed successfully: " + path);
			result = true;
		}
		else
		{
			Log.Warn("Store not found: " + path);
		}

		return result;
	}
}
