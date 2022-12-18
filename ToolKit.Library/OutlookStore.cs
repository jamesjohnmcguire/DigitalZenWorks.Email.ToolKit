/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookStore.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

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
		/// Remove all empty folders.
		/// </summary>
		/// <param name="store">The PST store to process.</param>
		/// <returns>The count of removed folders.</returns>
		public static int RemoveEmptyFolders(Store store)
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

					removedFolders = OutlookFolder.RemoveEmptyFolders(
						rootFolder, true);

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
		/// <param name="store">The PST store to process.</param>
		/// <returns>The count of removed folders.</returns>
		public static async Task<int> RemoveEmptyFoldersAsync(Store store)
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

					MAPIFolder rootFolder = store.GetRootFolder();

					removedFolders = await
						OutlookFolder.RemoveEmptyFoldersAsync(
							rootFolder, true).ConfigureAwait(false);

					Marshal.ReleaseComObject(rootFolder);
				}
			}

			Log.Info("Remove empty folder complete - total folders removed: " +
				removedFolders);

			return removedFolders;
		}

		/// <summary>
		/// Remove folder from PST store.
		/// </summary>
		/// <param name="path">The path of current folder.</param>
		/// <param name="subFolder">The sub-folder.</param>
		/// <param name="force">Whether to force the removal.</param>
		public static void RemoveFolder(
			string path,
			MAPIFolder subFolder,
			bool force)
		{
			if (subFolder != null)
			{
				string subFolderName = subFolder.Name;
				MAPIFolder parentFolder = OutlookFolder.GetParent(subFolder);

				if (parentFolder != null)
				{
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

					OutlookFolder.RemoveFolder(subFolder, index, force);
				}
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
		/// Get the item's synopses.
		/// </summary>
		/// <param name="entryId">The entryId of the MailItem to check.</param>
		/// <returns>The synoses of the item.</returns>
		public string GetItemSynopses(string entryId)
		{
			NameSpace session = outlookAccount.Session;
			MailItem mailItem = session.GetItemFromID(entryId);
			string synopses = MapiItem.GetItemSynopses(mailItem);

			return synopses;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <param name="entryId">The entryId of the MailItem to check.</param>
		/// <returns>The synoses of the item.</returns>
		public MailItem GetMailItemFromEntryId(string entryId)
		{
			NameSpace session = outlookAccount.Session;
			MailItem mailItem = session.GetItemFromID(entryId);

			return mailItem;
		}

		/// <summary>
		/// Get the total duplicates in the store.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		/// <returns>A list of total duplicates in the store.</returns>
		public IDictionary<string, IList<string>> GetTotalDuplicates(
			string pstFilePath)
		{
			IDictionary<string, IList<string>> hashTable =
				new Dictionary<string, IList<string>>();

			Store store = outlookAccount.GetStore(pstFilePath);

			if (store != null)
			{
				MAPIFolder rootFolder = store.GetRootFolder();
				OutlookFolder outlookFolder = new (outlookAccount);

				hashTable = outlookFolder.GetItemHashes(rootFolder);
			}

			return hashTable;
		}

		/// <summary>
		/// List the folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		/// <param name="folderPath">The folder path to check.</param>
		/// <param name="recurse">Indicates whether to recurse into
		/// sub-folders or not.</param>
		/// <returns>The folders.</returns>
		public IList<string> ListFolders(
			string pstFilePath, string folderPath, bool recurse)
		{
			IList<string> folderNames = new List<string>();

			Store store = outlookAccount.GetStore(pstFilePath);

			if (store != null)
			{
				MAPIFolder folder = OutlookFolder.CreateFolderPath(
					store, folderPath);

				if (folder != null)
				{
					folderNames = OutlookFolder.ListFolders(
						folderNames, folderPath, folder, recurse);

					Marshal.ReleaseComObject(folder);
				}
			}

			return folderNames;
		}

		/// <summary>
		/// List the top senders in  the store.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		/// <param name="amount">The amout of senders to list.</param>
		/// <returns>The top senders.</returns>
		public IList<KeyValuePair<string, int>> ListTopSenders(
			string pstFilePath, int amount)
		{
			IList<KeyValuePair<string, int>> topSenders =
				new List<KeyValuePair<string, int>>();

			Store store = outlookAccount.GetStore(pstFilePath);

			MAPIFolder rootFolder = store.GetRootFolder();

			string storePath = GetStoreName(store);
			storePath += "::";

			IDictionary<string, int> sendersCounts =
				new Dictionary<string, int>();

			sendersCounts = OutlookFolder.GetSendersCount(
				storePath, rootFolder, sendersCounts);

			Marshal.ReleaseComObject(rootFolder);
			Marshal.ReleaseComObject(store);

			IOrderedEnumerable<KeyValuePair<string, int>> orderedList =
				sendersCounts.OrderByDescending(pair => pair.Value);
			IEnumerable<KeyValuePair<string, int>> orderedListTop =
				orderedList.Take(amount);

			topSenders = orderedListTop.ToList();

			return topSenders;
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		public void MergeFolders(string pstFilePath, bool dryRun)
		{
			Store store = outlookAccount.GetStore(pstFilePath);

			MergeFolders(store, dryRun);

			Log.Info("Merge folders complete - total folders checked: " +
				totalFolders);
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="store">The store to check.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>The total folders checked.</returns>
		public uint MergeFolders(Store store, bool dryRun)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store);
				Log.Info("Merging folders in: " + storePath);

				storePath += "::";
				MAPIFolder rootFolder = store.GetRootFolder();

				OutlookFolder outlookFolder = new (outlookAccount);
				outlookFolder.MergeFolders(storePath, rootFolder, dryRun);

				totalFolders++;

				Marshal.ReleaseComObject(rootFolder);
				Marshal.ReleaseComObject(store);
			}

			return totalFolders;
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MergeFoldersAsync(string pstFilePath, bool dryRun)
		{
			Store store = outlookAccount.GetStore(pstFilePath);

			await MergeFoldersAsync(store, dryRun).ConfigureAwait(false);

			Log.Info("Merge folders complete - total folders checked: " +
				totalFolders);
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="store">The store to check.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>The total folders checked.</returns>
		public async Task<uint> MergeFoldersAsync(Store store, bool dryRun)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store);
				Log.Info("Merging folders in: " + storePath);

				storePath += "::";
				MAPIFolder rootFolder = store.GetRootFolder();

				OutlookFolder outlookFolder = new (outlookAccount);
				await outlookFolder.MergeFoldersAsync(
					storePath, rootFolder, dryRun).ConfigureAwait(false);

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

				LogFormatMessage.Info(
					"Moving contents of {0} to {1}",
					sourcePath,
					destinationPath);

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
							subFolder, destinationSubFolder);

						// Once all the items have been moved,
						// now remove the folder.
						bool isReserved =
							OutlookFolder.IsReservedFolder(subFolder);

						if (isReserved == false)
						{
							OutlookFolder.RemoveFolder(
								subFolder, subIndex, false);
						}
					}
					else
					{
						// Folder doesn't already exist, so just move it.
						LogFormatMessage.Info(
							"at: {0} Moving {1} to {2}",
							subPath,
							folderName,
							destinationPath);

						subFolder.MoveTo(destinationRootFolder);
					}
				}
			}
		}

		/// <summary>
		/// Merge 2 stores together.
		/// </summary>
		/// <param name="sourcePstPath">The source PST path.</param>
		/// <param name="destinationPstPath">The desination PST path.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MergeStoresAsync(
			string sourcePstPath, string destinationPstPath)
		{
			Store source = outlookAccount.GetStore(sourcePstPath);
			Store destination = outlookAccount.GetStore(destinationPstPath);

			await MergeStoresAsync(source, destination).ConfigureAwait(false);
		}

		/// <summary>
		/// Merge 2 stores together.
		/// </summary>
		/// <param name="source">The source store.</param>
		/// <param name="destination">The desination store.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MergeStoresAsync(Store source, Store destination)
		{
			if (source != null && destination != null)
			{
				string sourcePath = GetStoreName(source);
				string destinationPath = GetStoreName(destination);

				LogFormatMessage.Info(
					"Moving contents of {0} to {1}",
					sourcePath,
					destinationPath);

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

						await outlookFolder.MoveFolderContentsAsync(
							subFolder, destinationSubFolder).
								ConfigureAwait(false);

						// Once all the items have been moved,
						// now remove the folder.
						bool isReserved =
							OutlookFolder.IsReservedFolder(subFolder);

						if (isReserved == false)
						{
							OutlookFolder.RemoveFolder(
								subFolder, subIndex, false);
						}
					}
					else
					{
						// Folder doesn't already exist, so just move it.
						LogFormatMessage.Info(
							"at: {0} Moving {1} to {2}",
							subPath,
							folderName,
							destinationPath);

						subFolder.MoveTo(destinationRootFolder);
					}
				}
			}
		}

		/// <summary>
		/// Move folder.
		/// </summary>
		/// <param name="sourcePstPath">The source PST path.</param>
		/// <param name="sourceFolderPath">The source folder path.</param>
		/// <param name="destinationPstPath">The desination PST path.</param>
		/// <param name="destinationFolderPath">The destination folder path.</param>
		public void MoveFolder(
			string sourcePstPath,
			string sourceFolderPath,
			string destinationPstPath,
			string destinationFolderPath)
		{
			Store source = outlookAccount.GetStore(sourcePstPath);
			Store destination;

			if (string.IsNullOrWhiteSpace(destinationPstPath) ||
				destinationPstPath.Equals(
					sourcePstPath, StringComparison.OrdinalIgnoreCase))
			{
				destination = source;
			}
			else
			{
				destination = outlookAccount.GetStore(destinationPstPath);
			}

			MoveFolder(
				source, sourceFolderPath, destination, destinationFolderPath);
		}

		/// <summary>
		/// Move folder.
		/// </summary>
		/// <param name="source">The source store.</param>
		/// <param name="sourceFolderPath">The source folder path.</param>
		/// <param name="destination">The destination store.</param>
		/// <param name="destinationFolderPath">The destination folder path.</param>
		public void MoveFolder(
			Store source,
			string sourceFolderPath,
			Store destination,
			string destinationFolderPath)
		{
			if (source != null)
			{
				try
				{
					bool folderExists = OutlookFolder.DoesFolderExist(
						source, sourceFolderPath);

					// If source folder doesn't exist, there is nothing to do.
					if (folderExists == true)
					{
						MAPIFolder sourceFolder = OutlookFolder.CreateFolderPath(
							source, sourceFolderPath);

						MAPIFolder destinationParent = OutlookFolder.GetPathParent(
							destination, destinationFolderPath);

						folderExists = OutlookFolder.DoesFolderExist(
							destination, destinationFolderPath);

						string parentPath =
							OutlookFolder.GetFolderPath(destinationParent);
						string folderName = sourceFolder.Name;
						string destinationName =
							OutlookFolder.GetBaseFolderName(
								destinationFolderPath);

						if (folderExists == true)
						{
							MAPIFolder destinationFolder =
								OutlookFolder.CreateFolderPath(
									destination, destinationFolderPath);

							MoveExistingFolder(
								sourceFolder, destinationFolder, destinationName);
						}
						else
						{
							// Folder doesn't already exist, so just move it.
							LogFormatMessage.Info(
								"at: {0} Moving {1} to {2}",
								parentPath,
								folderName,
								destinationName);

							bool isRootFolder =
								OutlookFolder.IsRootFolder(sourceFolder);

							if (isRootFolder == true)
							{
								sourceFolder.MoveTo(destinationParent);
							}
							else
							{
								string destinationParentId =
									destinationParent.EntryID;

								MAPIFolder sourceParent = OutlookFolder.GetParent(sourceFolder);

								if (sourceParent != null)
								{
									string sourceParentId =
										sourceParent.EntryID;

									if (sourceParentId.Equals(
										destinationParentId,
										StringComparison.OrdinalIgnoreCase))
									{
										sourceFolder.Name = destinationName;
									}
									else
									{
										sourceFolder.MoveTo(destinationParent);

										if (!sourceFolder.Name.Equals(
											destinationName,
											StringComparison.OrdinalIgnoreCase))
										{
											sourceFolder.Name =
												destinationName;
										}
									}
								}
							}
						}
					}
				}
				catch (COMException exception)
				{
					Log.Error(exception.ToString());
				}
			}
		}

		/// <summary>
		/// Move folder.
		/// </summary>
		/// <param name="sourcePstPath">The source PST path.</param>
		/// <param name="sourceFolderPath">The source folder path.</param>
		/// <param name="destinationPstPath">The desination PST path.</param>
		/// <param name="destinationFolderPath">The destination folder path.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MoveFolderAsync(
			string sourcePstPath,
			string sourceFolderPath,
			string destinationPstPath,
			string destinationFolderPath)
		{
			Store source = outlookAccount.GetStore(sourcePstPath);
			Store destination;

			if (string.IsNullOrWhiteSpace(destinationPstPath) ||
				destinationPstPath.Equals(
					sourcePstPath, StringComparison.OrdinalIgnoreCase))
			{
				destination = source;
			}
			else
			{
				destination = outlookAccount.GetStore(destinationPstPath);
			}

			await MoveFolderAsync(
				source, sourceFolderPath, destination, destinationFolderPath).
					ConfigureAwait(false);
		}

		/// <summary>
		/// Move folder.
		/// </summary>
		/// <param name="source">The source store.</param>
		/// <param name="sourceFolderPath">The source folder path.</param>
		/// <param name="destination">The destination store.</param>
		/// <param name="destinationFolderPath">The destination folder path.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MoveFolderAsync(
			Store source,
			string sourceFolderPath,
			Store destination,
			string destinationFolderPath)
		{
			if (source != null)
			{
				try
				{
					bool folderExists = OutlookFolder.DoesFolderExist(
						source, sourceFolderPath);

					// If source folder doesn't exist, there is nothing to do.
					if (folderExists == true)
					{
						MAPIFolder sourceFolder =
							OutlookFolder.CreateFolderPath(
								source, sourceFolderPath);

						MAPIFolder destinationParent =
							OutlookFolder.GetPathParent(
								destination, destinationFolderPath);

						folderExists = OutlookFolder.DoesFolderExist(
							destination, destinationFolderPath);

						string parentPath =
							OutlookFolder.GetFolderPath(destinationParent);
						string folderName = sourceFolder.Name;
						string destinationName =
							OutlookFolder.GetBaseFolderName(
								destinationFolderPath);

						if (folderExists == true)
						{
							MAPIFolder destinationFolder =
								OutlookFolder.CreateFolderPath(
									destination, destinationFolderPath);

							await MoveExistingFolderAsync(
								sourceFolder,
								destinationFolder,
								destinationName).ConfigureAwait(false);
						}
						else
						{
							// Folder doesn't already exist, so just move it.
							LogFormatMessage.Info(
								"at: {0} Moving {1} to {2}",
								parentPath,
								folderName,
								destinationName);

							bool isRootFolder =
								OutlookFolder.IsRootFolder(sourceFolder);

							if (isRootFolder == true)
							{
								sourceFolder.MoveTo(destinationParent);
							}
							else
							{
								string destinationParentId =
									destinationParent.EntryID;

								MAPIFolder sourceParent =
									OutlookFolder.GetParent(sourceFolder);

								if (sourceParent != null)
								{
									string sourceParentId =
										sourceParent.EntryID;

									if (sourceParentId.Equals(
										destinationParentId,
										StringComparison.OrdinalIgnoreCase))
									{
										sourceFolder.Name = destinationName;
									}
									else
									{
										sourceFolder.MoveTo(destinationParent);

										if (!sourceFolder.Name.Equals(
											destinationName,
											StringComparison.OrdinalIgnoreCase))
										{
											sourceFolder.Name =
												destinationName;
										}
									}
								}
							}
						}
					}
				}
				catch (COMException exception)
				{
					Log.Error(exception.ToString());
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
				int removedDuplicates =
					outlookFolder.RemoveDuplicates(rootFolder, dryRun);

				if (flush == true)
				{
					EmptyDeletedItemsFolder(store);
				}

				LogFormatMessage.Info(
					"Duplicates Removed in: {0}: {1}",
					storePath,
					removedDuplicates.ToString(CultureInfo.InvariantCulture));

				totalFolders++;
				Marshal.ReleaseComObject(rootFolder);
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
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task RemoveDuplicatesAsync(
			string storePath, bool dryRun, bool flush)
		{
			Store store = outlookAccount.GetStore(storePath);

			if (store != null)
			{
				await RemoveDuplicatesAsync(store, dryRun, flush).
					ConfigureAwait(false);
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
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task RemoveDuplicatesAsync(
			Store store, bool dryRun, bool flush)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store);
				Log.Info("Checking for duplicates in: " + storePath);

				MAPIFolder rootFolder = store.GetRootFolder();

				OutlookFolder outlookFolder = new (outlookAccount);

				int removedDuplicates = await
					outlookFolder.RemoveDuplicatesAsync(
						rootFolder, dryRun).ConfigureAwait(false);

				if (flush == true)
				{
					EmptyDeletedItemsFolder(store);
				}

				LogFormatMessage.Info(
					"Duplicates Removed in: {0}: {1}",
					storePath,
					removedDuplicates.ToString(CultureInfo.InvariantCulture));

				totalFolders++;
				Marshal.ReleaseComObject(rootFolder);
			}
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		/// <returns>The count of removed folders.</returns>
		public int RemoveEmptyFolders(string pstFilePath)
		{
			Store store = outlookAccount.GetStore(pstFilePath);

			int removedFolders = RemoveEmptyFolders(store);

			return removedFolders;
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		/// <returns>The count of removed folders.</returns>
		public async Task<int> RemoveEmptyFoldersAsync(string pstFilePath)
		{
			Store store = outlookAccount.GetStore(pstFilePath);

			int removedFolders =
				await RemoveEmptyFoldersAsync(store).ConfigureAwait(false);

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
			Folders folders = deletedItemsFolder.Folders;
			int totalItems = folders.Count;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group will change.
			for (int index = totalItems; index > 0; index--)
			{
				MAPIFolder folder = folders[index];
				folder.Delete();
			}

			Items items = deletedItemsFolder.Items;
			totalItems = items.Count;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group will change.
			for (int index = totalItems; index > 0; index--)
			{
				object item = items[index];

				MapiItem.DeleteItem(item);
			}
		}

		private void MoveExistingFolder(
			MAPIFolder sourceFolder,
			MAPIFolder destinationFolder,
			string destinationName)
		{
			if (sourceFolder.EntryID.Equals(
				destinationFolder.EntryID, StringComparison.OrdinalIgnoreCase))
			{
				// Special case: If the names have different case-sensitivity,
				// rename to requested.
				if (sourceFolder.Name.Equals(
					destinationName,
					StringComparison.OrdinalIgnoreCase))
				{
					sourceFolder.Name = destinationName;
				}
				else
				{
					Log.Warn("Not moving folder to itself");
				}
			}
			else
			{
				OutlookFolder outlookFolder = new (outlookAccount);

				outlookFolder.MoveFolderContents(
					sourceFolder, destinationFolder);

				// Once all the items have been moved, remove the folder.
				OutlookFolder.SafeDelete(sourceFolder);
			}
		}

		private async Task MoveExistingFolderAsync(
			MAPIFolder sourceFolder,
			MAPIFolder destinationFolder,
			string destinationName)
		{
			if (sourceFolder.EntryID.Equals(
				destinationFolder.EntryID, StringComparison.OrdinalIgnoreCase))
			{
				// Special case: If the names have different case-sensitivity,
				// rename to requested.
				if (sourceFolder.Name.Equals(
						destinationName,
						StringComparison.OrdinalIgnoreCase) &&
					!sourceFolder.Name.Equals(
					destinationName,
					StringComparison.Ordinal))
				{
					sourceFolder.Name = destinationName;
				}
				else
				{
					Log.Warn("Not moving folder to itself");
				}
			}
			else
			{
				OutlookFolder outlookFolder = new (outlookAccount);

				await outlookFolder.MoveFolderContentsAsync(
					sourceFolder, destinationFolder).ConfigureAwait(false);

				// Once all the items have been moved, remove the folder.
				OutlookFolder.SafeDelete(sourceFolder);
			}
		}
	}
}
