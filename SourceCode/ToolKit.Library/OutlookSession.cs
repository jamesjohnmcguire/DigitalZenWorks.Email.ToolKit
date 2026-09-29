/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookSession.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.IO;
using System.Runtime.InteropServices;
using global::Common.Logging;
using Microsoft.Office.Interop.Outlook;

public class OutlookSession
	: IOutlookSession
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod() !.DeclaringType);

	private readonly NameSpace? session;

	public OutlookSession(Application application)
	{
		session = application.Session;
	}

	public OutlookSession(NameSpace? session)
	{
		this.session = session;
	}

	public OutlookFolder? GetFolderFromId(string entryId, string storeId)
	{
		OutlookFolder? folder = null;

		if (session != null)
		{
			MAPIFolder? mapiFolder = session.GetFolderFromID(entryId, storeId);

			if (mapiFolder != null)
			{
				folder = new(mapiFolder);
			}
		}

		return folder;
	}

	public object? GetItemFromId(string entryId)
	{
		object? item = null;

		if (session != null)
		{
			item = session.GetItemFromID(entryId);
		}

		return item;
	}

	/// <summary>
	/// Create a new pst storage file.
	/// </summary>
	/// <param name="path">The path to the pst file.</param>
	/// <returns>A store object or null if not found.</returns>
	public Store? GetStore(string path)
	{
		Store? store = null;

		if (session != null)
		{
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
				Store? checkStore = null;

				try
				{
					checkStore = session.Stores[index];
				}
				catch (UnauthorizedAccessException exception)
				{
					Log.Error(exception.ToString());
				}

				if (checkStore == null)
				{
					Log.Warn("Enumerating stores - store is null");
				}
				else
				{
					string filePath = checkStore.FilePath;

					if (!string.IsNullOrWhiteSpace(filePath) &&
						filePath.Equals(
							path, StringComparison.OrdinalIgnoreCase))
					{
						store = checkStore;
						break;
					}
				}
			}

			if (store == null)
			{
				Log.Warn("Store not found: " + path);
			}
		}

		return store;
	}

	public OutlookMail? OpenMailItemFile(string filePath)
	{
		OutlookMail? outlookMailItem = null;
		object? item = OpenSharedItem(filePath);

		if (item is MailItem mailItem)
		{
			outlookMailItem = new(mailItem);
		}
		else if (item is not null)
		{
			Marshal.ReleaseComObject(item);
		}

		return outlookMailItem;
	}

	public object? OpenSharedItem(string filePath)
	{
		object? item = null;

		if (session != null)
		{
			// session is Namespace
			item = session.OpenSharedItem(filePath);
		}

		return item;
	}

	/// <summary>
	/// Removes a store from Outlook.
	/// </summary>
	/// <param name="store">The store to remove.</param>
	/// <returns>remove result.</returns>
	public bool RemoveStore(Store store)
	{
		bool result = false;

		if (store != null)
		{
			Log.Info("Begin to Removing store: " + store.DisplayName);

			MAPIFolder rootFolder = store.GetRootFolder();

			if (session != null)
			{
				session.RemoveStore(rootFolder);

				Log.Info("Store removed successfully: " + store.DisplayName);
				result = true;
			}
		}
		else
		{
			Log.Warn("Store not present");
		}

		return result;
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

		Store? store = GetStore(path);

		if (store == null)
		{
			Log.Warn("Store not found: " + path);
		}
		else
		{
			result = RemoveStore(store);
		}

		return result;
	}

	internal MAPIFolder? GetFolderFromIdInternal(string entryId, string storeId)
	{
		MAPIFolder? mapiFolder = null;

		if (session != null)
		{
			mapiFolder = session.GetFolderFromID(entryId, storeId);
		}

		return mapiFolder;
	}
}
