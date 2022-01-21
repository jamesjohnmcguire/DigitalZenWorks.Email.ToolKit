/////////////////////////////////////////////////////////////////////////////
// <copyright file="PstOutlook.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Provides support for interating with an Outlook PST file.
	/// </summary>
	public class PstOutlook
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private readonly Application outlookApplication;
		private readonly NameSpace outlookNamespace;

		/// <summary>
		/// Initializes a new instance of the <see cref="PstOutlook"/> class.
		/// </summary>
		public PstOutlook()
		{
			outlookApplication = new ();

			outlookNamespace = outlookApplication.GetNamespace("mapi");
		}

		/// <summary>
		/// Add folder in safe context.
		/// </summary>
		/// <param name="parentFolder">The parent folder.</param>
		/// <param name="folderName">The new folder name.</param>
		/// <returns>The added folder.</returns>
		public static MAPIFolder AddFolderSafe(
			MAPIFolder parentFolder, string folderName)
		{
			MAPIFolder pstFolder = null;

			if (parentFolder != null)
			{
				Log.Info("Adding outlook folder: " + folderName);

				try
				{
					pstFolder =
						parentFolder.Folders.Add(folderName);
				}
				catch (COMException exception)
				{
					Log.Warn(exception.ToString());

					// Possibly already exists... ?
					try
					{
						pstFolder =
							parentFolder.Folders[folderName];
					}
					catch (COMException addionalException)
					{
						Log.Warn(addionalException.ToString());
					}
				}
			}

			return pstFolder;
		}

		/// <summary>
		/// Gets the message as a stream.
		/// </summary>
		/// <param name="filePath">The file path to create.</param>
		/// <returns>The message as a stream.</returns>
		public static Stream GetMsgFileStream(string filePath)
		{
			FileStream stream = new (filePath, FileMode.Create);

			return stream;
		}

		/// <summary>
		/// Get top level folder by name.
		/// </summary>
		/// <param name="pstStore">The pst store to check.</param>
		/// <param name="folderName">The folder name.</param>
		/// <returns>The MAPIFolder object.</returns>
		public static MAPIFolder GetTopLevelFolder(Store pstStore, string folderName)
		{
			MAPIFolder pstFolder = null;

			if (pstStore != null)
			{
				MAPIFolder rootFolder = pstStore.GetRootFolder();

				pstFolder = AddFolderSafe(rootFolder, folderName);
			}

			return pstFolder;
		}

		/// <summary>
		/// Add MSG file as MailItem in folder.
		/// </summary>
		/// <param name="pstFolder">The MSG file path.</param>
		/// <param name="filePath">The folder to add to.</param>
		public void AddMsgFile(MAPIFolder pstFolder, string filePath)
		{
			if (pstFolder != null && !string.IsNullOrWhiteSpace(filePath))
			{
				bool exists = File.Exists(filePath);

				if (exists == true)
				{
					MailItem item = outlookNamespace.OpenSharedItem(filePath);

					item.UnRead = false;
					item.Save();

					item.Move(pstFolder);

					Marshal.ReleaseComObject(item);
				}
				else
				{
					Log.Warn("File doesn't exist: " + filePath);
				}
			}
		}

		/// <summary>
		/// Create a new pst storage file.
		/// </summary>
		/// <param name="path">The path to the pst file.</param>
		/// <returns>A store object.</returns>
		public Store CreateStore(string path)
		{
			bool exists = File.Exists(path);

			if (exists == true)
			{
				Log.Warn("File already exists!: " + path);
			}

			Store newPst = null;

			// If the .pst file does not exist, Microsoft Outlook creates it.
			outlookNamespace.Session.AddStore(path);

			foreach (Store store in outlookNamespace.Session.Stores)
			{
				if (store.FilePath == path)
				{
					newPst = store;
					break;
				}
			}

			return newPst;
		}

		/// <summary>
		/// Gets folder from entry id.
		/// </summary>
		/// <param name="entryId">The entry id.</param>
		/// <param name="store">The store to check.</param>
		/// <returns>The folder.</returns>
		public MAPIFolder GetFolderFromID(string entryId, Store store)
		{
			MAPIFolder folder = null;

			if (store != null)
			{
				folder =
					outlookNamespace.GetFolderFromID(entryId, store.StoreID);
			}

			return folder;
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		public void RemoveEmptyFolders()
		{
			foreach (Store store in outlookNamespace.Session.Stores)
			{
				MAPIFolder rootFolder = store.GetRootFolder();

				for (int index = rootFolder.Folders.Count - 1;
					index >= 0; index--)
				{
					// Office uses 1 based indexes from VBA.
					int offset = index + 1;

					MAPIFolder subFolder = rootFolder.Folders[offset];
					bool subFolderEmtpy = RemoveEmptyFolders(subFolder);

					if (subFolderEmtpy == true)
					{
						Log.Info("Removing empty folder: " + subFolder.Name);
						rootFolder.Folders.Remove(offset);
					}
				}
			}
		}

		/// <summary>
		/// Create a new pst storage file.
		/// </summary>
		/// <param name="store">The store to check.</param>
		public void RemoveStore(Store store)
		{
			if (store != null)
			{
				MAPIFolder rootFolder = store.GetRootFolder();

				outlookNamespace.Session.RemoveStore(rootFolder);
			}
		}

		private bool RemoveEmptyFolders(MAPIFolder folder)
		{
			bool isEmpty = false;

			for (int index = folder.Folders.Count - 1; index >= 0; index--)
			{
				// Office uses 1 based indexes from VBA.
				int offset = index + 1;

				MAPIFolder subFolder = folder.Folders[offset];

				bool subFolderEmtpy = RemoveEmptyFolders(subFolder);

				if (subFolderEmtpy == true)
				{
					Log.Info("Removing empty folder: " + subFolder.Name);
					folder.Folders.Remove(offset);
				}
			}

			if (folder.Folders.Count == 0 && folder.Items.Count == 0)
			{
				isEmpty = true;
			}

			return isEmpty;
		}
	}
}
