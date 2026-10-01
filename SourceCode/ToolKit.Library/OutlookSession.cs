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

/// <summary>
/// Represents an Outlook session and provides methods to interact with Outlook
/// folders, items, and stores.
/// </summary>
public class OutlookSession
	: IOutlookSession
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod() !.DeclaringType);

	private readonly NameSpace? session;

	/// <summary>
	/// Initializes a new instance of the <see cref="OutlookSession"/> class.
	/// </summary>
	/// <param name="session">The Outlook session.</param>
	public OutlookSession(NameSpace? session)
	{
		this.session = session;
	}

	/// <summary>
	/// Gets a folder from its entry id and store id.
	/// </summary>
	/// <param name="entryId">The entry id of the folder.</param>
	/// <param name="storeId">The store id of the folder.</param>
	/// <returns>The folder if found; otherwise, null.</returns>
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

	/// <summary>
	/// Gets an item from its entry id.
	/// </summary>
	/// <param name="entryId">The entry id of the item.</param>
	/// <returns>The item if found; otherwise, null.</returns>
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
	/// <param name="createIfMissing">Whether to create the pst file if it does
	/// not exist.</param>
	/// <returns>A store object or null if not found.</returns>
	public Store? GetStore(string path, bool createIfMissing = true)
	{
		Store? store = null;

		if (session != null)
		{
			path = Path.GetFullPath(path);
			bool exists = File.Exists(path);

			if (exists == false && createIfMissing == false)
			{
				Log.Warn("Store file does not exist: " + path);
			}
			else
			{
				if (createIfMissing == true)
				{
					if (exists == false)
					{
						Log.Info("Attempting to create store file: " + path);
					}

					string extension = Path.GetExtension(path);

					if (!extension.Equals(
						".pst", StringComparison.OrdinalIgnoreCase))
					{
						// Attempt to fix mistaken or missing file extension.
						path += ".pst";
					}

					// If the .pst file does not exist, Outlook creates it.
					session.AddStore(path);
				}

				store = FindStore(path);
			}
		}

		return store;
	}

	/// <summary>
	/// Opens a mail item file and returns an OutlookMail object.
	/// </summary>
	/// <param name="filePath">The path to the mail item file.</param>
	/// <returns>The OutlookMail object or null if not found.</returns>
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

	/// <summary>
	/// Opens a shared item in Outlook.
	/// </summary>
	/// <param name="filePath">The path to the shared item.</param>
	/// <returns>The opened item or null if not found.</returns>
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
			string displayName = store.DisplayName;
			Log.Info("Begin to Removing store: " + displayName);

			MAPIFolder rootFolder = store.GetRootFolder();

			try
			{
				if (session != null)
				{
					session.RemoveStore(rootFolder);

					Log.Info("Store removed successfully: " + displayName);
					result = true;
				}
			}
			finally
			{
				Marshal.ReleaseComObject(rootFolder);
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
			try
			{
				result = RemoveStore(store);
			}
			finally
			{
				Marshal.ReleaseComObject(store);
			}
		}

		return result;
	}

	/// <summary>
	/// Gets a folder from its entry id and store id.
	/// </summary>
	/// <param name="entryId">The entry id of the folder.</param>
	/// <param name="storeId">The store id of the folder.</param>
	/// <returns>The MAPI folder if found; otherwise, null.</returns>
	internal MAPIFolder? GetFolderFromIdInternal(string entryId, string storeId)
	{
		MAPIFolder? mapiFolder = null;

		if (session != null)
		{
			mapiFolder = session.GetFolderFromID(entryId, storeId);
		}

		return mapiFolder;
	}

	private Store? FindStore(string storePath)
	{
		Store? store = null;

		if (session != null)
		{
			storePath = Path.GetFullPath(storePath);
			Stores stores = session.Stores;

			try
			{
				int total = stores.Count;

				for (int index = 1; index <= total; index++)
				{
					Store? checkStore = null;

					try
					{
						checkStore = stores[index];
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
						try
						{
							string filePath = checkStore.FilePath;

							if (!string.IsNullOrWhiteSpace(filePath) &&
								filePath.Equals(
									storePath, StringComparison.OrdinalIgnoreCase))
							{
								// Transfer this reference to the caller.
								store = checkStore;
								checkStore = null;
								break;
							}
						}
						finally
						{
							if (checkStore != null)
							{
								Marshal.ReleaseComObject(checkStore);
							}
						}
					}
				}
			}
			finally
			{
				Marshal.ReleaseComObject(stores);
			}

			if (store == null)
			{
				Log.Warn("Store not found: " + storePath);
			}
		}

		return store;
	}
}
