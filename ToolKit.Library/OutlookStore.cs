/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookStore.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Provides support for interacting with an Outlook PST file.
	/// </summary>
	public class OutlookStore
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private readonly OutlookAccount outlookAccount;

		private uint totalFolders;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookStore"/> class.
		/// </summary>
		/// <param name="outlookAccount">The outlook account object.</param>
		public OutlookStore(OutlookAccount outlookAccount)
		{
			this.outlookAccount = outlookAccount;
		}

		/// <summary>
		/// Gets or sets the total folders.
		/// </summary>
		/// <value>The total folders.</value>
		public uint TotalFolders
		{
			get { return totalFolders; }
			set { totalFolders = value; }
		}

		/// <summary>
		/// Empty deleted items folder.
		/// </summary>
		/// <param name="store">The store to access.</param>
		public static void EmptyDeletedItemsFolder(Store store)
		{
			if (store != null)
			{
				MAPIFolder deletedItemsFolder = store.GetDefaultFolder(
						OlDefaultFolders.olFolderDeletedItems);

				EmptyDeletedItemsFolder(deletedItemsFolder);
			}
		}

		/// <summary>
		/// Get store name.
		/// </summary>
		/// <param name="store">The store to access.</param>
		/// <returns>The store name.</returns>
		public static string GetStoreName(Store store)
		{
			string name = null;

			if (store != null)
			{
				name = store.DisplayName;

				if (string.IsNullOrWhiteSpace(name))
				{
					string path = store.FilePath;
					name = Path.GetFileNameWithoutExtension(path);
				}
			}

			return name;
		}

		/// <summary>
		/// Get top level folder by name.
		/// </summary>
		/// <param name="store">The store to check.</param>
		/// <param name="folderName">The folder name.</param>
		/// <returns>The MAPIFolder object.</returns>
		public static MAPIFolder GetTopLevelFolder(
			Store store, string folderName)
		{
			MAPIFolder pstFolder = null;

			if (store != null)
			{
				MAPIFolder rootFolder = store.GetRootFolder();

				pstFolder = OutlookFolder.AddFolder(rootFolder, folderName);

				Marshal.ReleaseComObject(rootFolder);
			}

			return pstFolder;
		}

		/// <summary>
		/// Remove folder from PST store.
		/// </summary>
		/// <param name="path">The path of current folder.</param>
		/// <param name="subFolder">The sub-folder.</param>
		/// <param name="force">Whether to force the removal.</param>
		public void RemoveFolder(
			string path,
			MAPIFolder subFolder,
			bool force)
		{
			if (subFolder != null)
			{
				string subFolderName = subFolder.Name;
				MAPIFolder parentFolder = subFolder.Parent;

				int count = parentFolder.Folders.Count;
				int index;
				for (index = 1; index <= count; index++)
				{
					MAPIFolder folder = parentFolder.Folders[index];

					string name = folder.Name;

					if (name.Equals(
						subFolderName, StringComparison.OrdinalIgnoreCase))
					{
						break;
					}
				}

				path += "/" + subFolder.Name;

				OutlookFolder outlookFolder = new (outlookAccount);
				outlookFolder.RemoveFolder(path, index, subFolder, force);
			}
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
				NameSpace session = outlookAccount.Session;
				folder = session.GetFolderFromID(entryId, store.StoreID);
			}

			return folder;
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		public void MergeFolders(string pstFilePath)
		{
			Store store = outlookAccount.GetStore(pstFilePath);

			MergeFolders(store);

			Log.Info("Merge folders complete - total folders checked: " +
				totalFolders);
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="store">The store to check.</param>
		/// <returns>The total folders checked.</returns>
		public uint MergeFolders(Store store)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store);
				Log.Info("Merging folders in: " + storePath);

				storePath += "::";
				MAPIFolder rootFolder = store.GetRootFolder();

				OutlookFolder outlookFolder = new (outlookAccount);
				outlookFolder.MergeFolders(storePath, rootFolder);

				totalFolders++;

				Marshal.ReleaseComObject(rootFolder);
				Marshal.ReleaseComObject(store);
			}

			return totalFolders;
		}

		/// <summary>
		/// Merge 2 stores together.
		/// </summary>
		/// <param name="sourcePstPath">The source PST path.</param>
		/// <param name="destinationPstPath">The desination PST path.</param>
		public void MergeStores(
			string sourcePstPath, string destinationPstPath)
		{
			Store source = outlookAccount.GetStore(sourcePstPath);
			Store destination = outlookAccount.GetStore(destinationPstPath);

			MergeStores(source, destination);
		}

		/// <summary>
		/// Merge 2 stores together.
		/// </summary>
		/// <param name="source">The source store.</param>
		/// <param name="destination">The desination store.</param>
		public void MergeStores(Store source, Store destination)
		{
			if (source != null && destination != null)
			{
				string sourcePath = GetStoreName(source);
				string destinationPath = GetStoreName(destination);

				string message = string.Format(
					CultureInfo.InvariantCulture,
					"Moving contents of {0} to {1}",
					sourcePath,
					destinationPath);
				Log.Info(message);

				MAPIFolder sourceRootFolder = source.GetRootFolder();
				MAPIFolder destinationRootFolder = destination.GetRootFolder();

				int subFolderCount = sourceRootFolder.Folders.Count;

				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group may change.
				for (int subIndex = subFolderCount; subIndex > 0; subIndex--)
				{
					MAPIFolder subFolder = sourceRootFolder.Folders[subIndex];
					string folderName = subFolder.Name;

					string subPath =
						destinationPath + "/" + folderName;

					bool folderExists = OutlookFolder.DoesFolderExist(
						destinationRootFolder, folderName);

					if (folderExists == true)
					{
						// Folder exists, so if just moving it, it will get
						// renamed something FolderName (2), so need to merge.
						MAPIFolder destinationSubFolder =
							OutlookFolder.GetSubFolder(
								destinationRootFolder, folderName);

						OutlookFolder outlookFolder = new (outlookAccount);

						outlookFolder.MoveFolderContents(
							subPath, subFolder, destinationSubFolder);

						// Once all the items have been moved,
						// now remove the folder.
						bool isReserved =
							OutlookFolder.IsReservedFolder(subFolder);

						if (isReserved == false)
						{
							outlookFolder.RemoveFolder(
								subPath, subIndex, subFolder, false);
						}
					}
					else
					{
						// Folder doesn't already exist, so just move it.
						message = string.Format(
							CultureInfo.InvariantCulture,
							"at: {0} Moving {1} to {2}",
							subPath,
							folderName,
							destinationPath);
						Log.Info(message);

						subFolder.MoveTo(destinationRootFolder);
					}
				}
			}
		}

		/// <summary>
		/// Remove duplicates items from the given store.
		/// </summary>
		/// <param name="storePath">The path of the PST file to
		/// process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <param name="flush">Indicates whether to empty the deleted items
		/// folder.</param>
		public void RemoveDuplicates(string storePath, bool dryRun, bool flush)
		{
			Store store = outlookAccount.GetStore(storePath);

			if (store != null)
			{
				RemoveDuplicates(store, dryRun, flush);
				Marshal.ReleaseComObject(store);
			}
		}

		/// <summary>
		/// Remove duplicates items from the given store.
		/// </summary>
		/// <param name="store">The PST store to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <param name="flush">Indicates whether to empty the deleted items
		/// folder.</param>
		public void RemoveDuplicates(Store store, bool dryRun, bool flush)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store);
				Log.Info("Checking for duplicates in: " + storePath);

				MAPIFolder rootFolder = store.GetRootFolder();

				OutlookFolder outlookFolder = new (outlookAccount);
				int[] duplicateCounts =
					outlookFolder.RemoveDuplicatesFromSubFolders(
						storePath, rootFolder, dryRun);

				if (flush == true)
				{
					EmptyDeletedItemsFolder(store);
				}

				int removedDuplicates =
					duplicateCounts[1] - duplicateCounts[0];
				string message = string.Format(
					CultureInfo.InvariantCulture,
					"Duplicates Removed in: {0}: {1}",
					storePath,
					removedDuplicates.ToString(CultureInfo.InvariantCulture));
				Log.Info(message);

				totalFolders++;
				Marshal.ReleaseComObject(rootFolder);
			}
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		public void RemoveEmptyFolders(string pstFilePath)
		{
			Store store = outlookAccount.GetStore(pstFilePath);

			RemoveEmptyFolders(store);
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="store">The PST store to process.</param>
		/// <returns>The count of removed folders.</returns>
		public int RemoveEmptyFolders(Store store)
		{
			int removedFolders = 0;

			if (store != null)
			{
				string path = store.FilePath;
				string extension = Path.GetExtension(path);

				if (!extension.Equals(
					".ost", StringComparison.OrdinalIgnoreCase))
				{
					string storePath = GetStoreName(store);

					Log.Info("Checking for empty folders in: " +
						storePath);
					storePath += "::";

					MAPIFolder rootFolder = store.GetRootFolder();

					removedFolders =
						RemoveEmptyFolders(storePath, rootFolder, 1);

					Marshal.ReleaseComObject(rootFolder);
				}
			}

			Log.Info("Remove empty folder complete - total folders removed: " +
				removedFolders);

			return removedFolders;
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="folder">The current folder.</param>
		/// <param name="index">The index of the folder.</param>
		/// <returns>The count of removed folders.</returns>
		public int RemoveEmptyFolders(
			string path, MAPIFolder folder, int index)
		{
			int removedFolders = 0;

			if (folder != null)
			{
				int subFolderCount = folder.Folders.Count;

				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group may change.
				for (int subIndex = subFolderCount; subIndex > 0; subIndex--)
				{
					MAPIFolder subFolder = folder.Folders[subIndex];

					string subPath = path + "/" + subFolder.Name;

					removedFolders +=
						RemoveEmptyFolders(subPath, subFolder, subIndex);

					Marshal.ReleaseComObject(subFolder);
				}

				if (folder.Folders.Count == 0 && folder.Items.Count == 0)
				{
					bool isReservedFolder =
						OutlookFolder.IsReservedFolder(folder);

					if (isReservedFolder == true)
					{
						string name = folder.Name;
						Log.Warn("Not deleting reserved folder: " +
							name);
					}
					else
					{
						OutlookFolder outlookFolder = new (outlookAccount);
						outlookFolder.RemoveFolder(
							path, index, folder, false);
						removedFolders++;
					}
				}
			}

			return removedFolders;
		}

		/// <summary>
		/// Create a new pst storage file.
		/// </summary>
		/// <param name="store">The store to check.</param>
		public void RemoveStore(Store store)
		{
			if (store != null)
			{
				NameSpace session = outlookAccount.Session;

				MAPIFolder rootFolder = store.GetRootFolder();

				session.RemoveStore(rootFolder);

				Marshal.ReleaseComObject(rootFolder);
				Marshal.ReleaseComObject(store);
			}
		}

		private static void EmptyDeletedItemsFolder(
			MAPIFolder deletedItemsFolder)
		{
			Items items = deletedItemsFolder.Items;
			int totalItems = items.Count;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group will change.
			for (int index = totalItems; index > 0; index--)
			{
				object item = items[index];

				MapiItem.DeleteItem(item);
			}
		}
	}
}
