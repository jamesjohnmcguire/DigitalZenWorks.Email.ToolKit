/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFolder.cs" company="James John McGuire">
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
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Delegate for a folder.
	/// </summary>
	/// <param name="folder">The folder to act upon.</param>
	/// <param name="conditional">A conditional clause to use within
	/// the delegate.</param>
	/// <returns>A value processed from the delegate.</returns>
	public delegate int FolderAction(MAPIFolder folder, bool conditional);

	/// <summary>
	/// Delegate for a folder.
	/// </summary>
	/// <param name="folder">The folder to act upon.</param>
	/// <param name="conditional">A conditional clause to use within
	/// the delegate.</param>
	/// <returns>A value processed from the delegate.</returns>
	public delegate Task<int> FolderActionAsync(
		MAPIFolder folder, bool conditional);

	/// <summary>
	/// Item action Delegate.
	/// </summary>
	/// <param name="folder">The folder to use.</param>
	/// <param name="item">The item to use.</param>
	public delegate void ItemAction(
		MAPIFolder folder, object item);

	/// <summary>
	/// Item action Delegate.
	/// </summary>
	/// <param name="folder">The folder to use.</param>
	/// <param name="item">The item to use.</param>
	/// <returns>A <see cref="Task"/> representing the asynchronous
	/// operation.</returns>
	public delegate Task ItemActionAsync(
		MAPIFolder folder, object item);

	/// <summary>
	/// Represents an Outlook Folder.
	/// </summary>
	public class OutlookFolder
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private static readonly string[] DeletedFolders =
		{
			"Deleted Items", "Deleted Messages"
		};

		private static readonly string[] ReservedFolders =
		{
			"Calendar", "Contacts", "Conversation Action Settings",
			"Deleted Items", "Deleted Messages", "Drafts", "Inbox",
			"Junk E-mail", "Journal", "Notes", "Outbox", "Quick Step Settings",
			"RSS Feeds", "Search Folders", "Sent Items", "Tasks"
		};

		private readonly OutlookAccount outlookAccount;

		private IDictionary<string, int> sendersCounts =
			new Dictionary<string, int>();

		private IDictionary<string, IList<string>> storeHashTable =
			new Dictionary<string, IList<string>>();

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookFolder"/> class.
		/// </summary>
		/// <param name="outlookAccount">The outlook account object.</param>
		public OutlookFolder(OutlookAccount outlookAccount)
		{
			this.outlookAccount = outlookAccount;
		}

		/// <summary>
		/// Gets or sets the amount of removed duplicates.
		/// </summary>
		/// <value>The amount of removed duplicates.</value>
		public int RemovedDuplicates { get; set; }

		/// <summary>
		/// Gets or sets the amount of duplicate sets.
		/// </summary>
		/// <remarks>A duplicate set is the set of emails that have the same
		/// hash signature.</remarks>
		/// <value>The amount of duplicate sets.</value>
		public int DuplicatesSets { get; set; }

		/// <summary>
		/// Add folder in safe context.
		/// </summary>
		/// <remarks>If there is a folder already existing with the given
		/// folder name, this method will return that folder.</remarks>
		/// <param name="parentFolder">The parent folder.</param>
		/// <param name="folderName">The new folder name.</param>
		/// <returns>The added or existing folder.</returns>
		public static MAPIFolder AddFolder(
			MAPIFolder parentFolder, string folderName)
		{
			MAPIFolder pstFolder = null;

			if (parentFolder != null && !string.IsNullOrWhiteSpace(folderName))
			{
				pstFolder = GetSubFolder(parentFolder, folderName);

				if (pstFolder == null)
				{
					string parentPath = GetFolderPath(parentFolder);
					Log.Info("At: " + parentPath + " Adding outlook folder: " +
						folderName);

					try
					{
						pstFolder = parentFolder.Folders.Add(folderName);
					}
					catch (COMException exception)
					{
						Log.Warn(exception.ToString());
					}
				}
			}

			return pstFolder;
		}

		/// <summary>
		/// Create folder path.
		/// </summary>
		/// <param name="store">The PST file store to use.</param>
		/// <param name="path">The full path to create.</param>
		/// <returns>The folder with the full path.</returns>
		public static MAPIFolder CreateFolderPath(Store store, string path)
		{
			MAPIFolder currentFolder = GetPathFolder(store, path, false);

			return currentFolder;
		}

		/// <summary>
		/// Does folder exist.
		/// </summary>
		/// <param name="parentFolder">The parent folder to check.</param>
		/// <param name="folderName">The name of the sub-folder.</param>
		/// <returns>Indicates whether the folder exists.</returns>
		public static bool DoesFolderExist(
			MAPIFolder parentFolder, string folderName)
		{
			bool folderExists = false;

			MAPIFolder folder = GetSubFolder(parentFolder, folderName);

			if (folder != null)
			{
				folderExists = true;
			}

			return folderExists;
		}

		/// <summary>
		/// Does folder exist.
		/// </summary>
		/// <param name="store">The PST file store to use.</param>
		/// <param name="path">The full path of the folder to check.</param>
		/// <returns>Indicates whether the folder exists.</returns>
		public static bool DoesFolderExist(Store store, string path)
		{
			bool folderExists = false;

			if (store != null && !string.IsNullOrWhiteSpace(path))
			{
				MAPIFolder currentFolder = store.GetRootFolder();

				string[] parts = GetPathParts(path);

				folderExists = true;

				for (int index = 0; index < parts.Length; index++)
				{
					string part = parts[index];

					if (index == 0)
					{
						string rootFolderName = currentFolder.Name;

						if (part.Equals(
							rootFolderName,
							StringComparison.OrdinalIgnoreCase))
						{
							// root, so skip over
							continue;
						}
					}

					currentFolder = GetSubFolder(currentFolder, part, false);

					if (currentFolder == null)
					{
						folderExists = false;
						break;
					}
				}
			}

			return folderExists;
		}

		/// <summary>
		/// Get the base folder name.
		/// </summary>
		/// <param name="folderPath">The folder path to check.</param>
		/// <returns>The base folder name.</returns>
		public static string GetBaseFolderName(string folderPath)
		{
			string folderName = null;

			if (!string.IsNullOrEmpty(folderPath))
			{
				string[] parts = GetPathParts(folderPath);
				folderName = parts[^1];
			}

			return folderName;
		}

		/// <summary>
		/// Get the folder's full path.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>The folder's full path.</returns>
		public static string GetFolderPath(MAPIFolder folder)
		{
			string path = null;

			if (folder != null)
			{
				string storeName = OutlookStore.GetStoreName(folder.Store);
				path = folder.Name;

				do
				{
					folder = GetParent(folder);

					if (folder != null)
					{
						string name = folder.Name;
						path = name + "/" + path;
					}
				}
				while (folder != null);

				path = storeName + "::" + path;
			}

			return path;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <param name="mailItem">The MailItem to check.</param>
		/// <returns>The synoses of the item.</returns>
		[Obsolete("GetMailItemSynopses is deprecated, " +
			"please use MapiItem.GetItemSynopses instead.")]
		public static string GetMailItemSynopses(MailItem mailItem)
		{
			string synopses = MapiItem.GetItemSynopses(mailItem);

			return synopses;
		}

		/// <summary>
		/// Get the current folder's parent.
		/// </summary>
		/// <param name="folder">The current folder.</param>
		/// <returns>The parent folder.</returns>
		public static MAPIFolder GetParent(MAPIFolder folder)
		{
			MAPIFolder parent = null;

			if (folder != null && folder.Parent is not null &&
				folder.Parent is MAPIFolder)
			{
				parent = folder.Parent;
			}

			return parent;
		}

		/// <summary>
		/// Get the parent folder of the given path.
		/// </summary>
		/// <param name="store">The PST file store to use.</param>
		/// <param name="folderPath">The folder path to check.</param>
		/// <returns>The parent folder of the given path.</returns>
		public static MAPIFolder GetPathParent(Store store, string folderPath)
		{
			MAPIFolder parent = GetPathFolder(store, folderPath, true);

			return parent;
		}

		/// <summary>
		/// Get senders counts.
		/// </summary>
		/// <param name="path">The current folder path.</param>
		/// <param name="folder">The folder to check.</param>
		/// <param name="sendersCounts">The current counts of senders.</param>
		/// <returns>The count of each sender.</returns>
		[Obsolete("GetSendersCount(string, int, MAPIFolder, " +
			"IDictionary<string, int>) is deprecated, " +
			"please use GetSendersCount(MAPIFolder) instead.")]
		public static IDictionary<string, int> GetSendersCount(
			string path,
			MAPIFolder folder,
			IDictionary<string, int> sendersCounts)
		{
			if (folder != null && sendersCounts != null)
			{
				OutlookAccount outlookAccount = OutlookAccount.Instance;
				OutlookFolder outlookFolder = new (outlookAccount);

				sendersCounts = outlookFolder.GetSendersCount(folder);
			}

			return sendersCounts;
		}

		/// <summary>
		/// Get sub folder from parent.
		/// </summary>
		/// <param name="parentFolder">The parent folder.</param>
		/// <param name="folderName">The new folder name.</param>
		/// <param name="caseSensitive">Indicates whether the check should
		/// be case-sensitive.</param>
		/// <returns>The added folder.</returns>
		public static MAPIFolder GetSubFolder(
			MAPIFolder parentFolder,
			string folderName,
			bool caseSensitive = false)
		{
			MAPIFolder pstFolder = null;

			if (parentFolder != null && !string.IsNullOrWhiteSpace(folderName))
			{
				if (caseSensitive == false)
				{
					try
					{
						pstFolder = parentFolder.Folders[folderName];
					}
					catch (COMException)
					{
					}
				}
				else
				{
					int total = parentFolder.Folders.Count;

					for (int index = 1; index <= total; index++)
					{
						MAPIFolder subFolder = parentFolder.Folders[index];

						if (folderName.Equals(
							subFolder.Name, StringComparison.Ordinal))
						{
							pstFolder = subFolder;
							break;
						}

						Marshal.ReleaseComObject(subFolder);
					}
				}
			}

			return pstFolder;
		}

		/// <summary>
		/// Is deleted folder.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>Indicates whether this is an system deleted items
		/// folder.</returns>
		public static bool IsDeletedFolder(MAPIFolder folder)
		{
			bool isDeletedFolder = false;

			if (folder != null)
			{
				string name = folder.Name;

				if (DeletedFolders.Contains(name))
				{
					// Only top level folders are reserved
					MAPIFolder parent = GetParent(folder);

					if (parent != null)
					{
						// Check if root folder
						bool isRoot = IsRootFolder(parent);

						if (isRoot == true)
						{
							isDeletedFolder = true;
						}

						Marshal.ReleaseComObject(parent);
					}
				}
			}

			return isDeletedFolder;
		}

		/// <summary>
		/// Indicates whether the given folder is a reserved folder.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>A value that indicates whether the given folder is a
		/// reserved folder.</returns>
		public static bool IsReservedFolder(MAPIFolder folder)
		{
			bool reserved = false;

			if (folder != null)
			{
				string name = folder.Name;

				if (ReservedFolders.Contains(name))
				{
					// Only top level folders are reserved
					MAPIFolder parent = GetParent(folder);

					if (parent != null)
					{
						bool isRoot = IsRootFolder(parent);

						if (isRoot == true)
						{
							reserved = true;
						}

						Marshal.ReleaseComObject(parent);
					}
				}
				else
				{
					bool isRoot = IsRootFolder(folder);

					if (isRoot == true)
					{
						reserved = true;
					}
				}
			}

			return reserved;
		}

		/// <summary>
		/// Indicates whether the given folder is the root folder.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>A value that indicates whether the given folder is the
		/// root folder.</returns>
		public static bool IsRootFolder(MAPIFolder folder)
		{
			bool isRootFolder = false;

			if (folder != null)
			{
				if (folder.Parent is null || folder.Parent is not MAPIFolder)
				{
					isRootFolder = true;
				}
			}

			return isRootFolder;
		}

		/// <summary>
		/// Indicates whether the given folder is a top level folder.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>A value that indicates whether the given folder is a
		/// top level folder.</returns>
		public static bool IsTopLevelFolder(MAPIFolder folder)
		{
			bool topLevel = false;

			if (folder != null)
			{
				MAPIFolder parent = GetParent(folder);

				if (parent != null)
				{
					bool isRoot = IsRootFolder(parent);

					if (isRoot == true)
					{
						topLevel = true;
					}

					Marshal.ReleaseComObject(parent);
				}
				else
				{
					bool isRoot = IsRootFolder(folder);

					if (isRoot == true)
					{
						// Also, include the root
						topLevel = true;
					}
				}
			}

			return topLevel;
		}

		/// <summary>
		/// List the folders.
		/// </summary>
		/// <param name="folderNames">The current list of folder names.</param>
		/// <param name="folderPath">The folder path to check.</param>
		/// <param name="folder">The folder to act upon.</param>
		/// <param name="recurse">Indicates whether to recurse into
		/// sub-folders or not.</param>
		/// <returns>The folders.</returns>
		public static IList<string> ListFolders(
			IList<string> folderNames,
			string folderPath,
			MAPIFolder folder,
			bool recurse)
		{
			if (folderNames != null && folder != null)
			{
				int count = folder.Folders.Count;
				for (int index = 1; index <= count; index++)
				{
					MAPIFolder subFolder = folder.Folders[index];

					string subFolderName = subFolder.Name;

					if (recurse == true)
					{
						string subFolderPath =
							folderPath + "/" + subFolderName;
						folderNames.Add(subFolderPath);

						folderNames = ListFolders(
							folderNames, subFolderPath, subFolder, true);
					}
					else
					{
						folderNames.Add(subFolderName);
					}

					Marshal.ReleaseComObject(subFolder);
				}

				Marshal.ReleaseComObject(folder);
			}

			return folderNames;
		}

		/// <summary>
		/// Normalize the folder name.
		/// </summary>
		/// <param name="folderName">The name of the folder to check.</param>
		/// <returns>The new folder name.</returns>
		/// <remarks>The returned folder name may often be the same as
		/// the given parameter.</remarks>
		public static string NormalizeFolderName(string folderName)
		{
			string duplicatePattern = CheckFolderNameNormalization(folderName);

			if (!string.IsNullOrWhiteSpace(duplicatePattern))
			{
				folderName =
					GetNormalizedFolderName(folderName, duplicatePattern);
			}

			return folderName;
		}

		/// <summary>
		/// Normalize folder path with forward slashes.
		/// </summary>
		/// <param name="path">The folder path.</param>
		/// <returns>The normalized folder path.</returns>
		public static string NormalizePath(string path)
		{
			if (!string.IsNullOrWhiteSpace(path))
			{
				path = path.Replace('\\', '/');
			}

			return path;
		}

		/// <summary>
		/// Recurse folders.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <param name="condition">A conditional to check.</param>
		/// <param name="folderAction">The delegate to act upon.</param>
		/// <returns>A value processed from the delegate.</returns>
		public static int RecurseFolders(
			MAPIFolder folder,
			bool condition,
			FolderAction folderAction)
		{
			int processed = 0;

			if (folder != null && folderAction != null)
			{
				try
				{
					bool isDeletedFolder = IsDeletedFolder(folder);

					// Skip processing of system deleted items folder.
					if (isDeletedFolder == false)
					{
						int folderCount = folder.Folders.Count;

						// Office uses 1 based indexes from VBA.
						// Iterate in reverse order as the group may change.
						for (int index = folderCount; index > 0; index--)
						{
							try
							{
								MAPIFolder subFolder = folder.Folders[index];

								processed += RecurseFolders(
									subFolder, condition, folderAction);

								Marshal.ReleaseComObject(subFolder);
							}
							catch (COMException exception)
							{
								string path = GetFolderPath(folder);

								string message = string.Format(
									CultureInfo.InvariantCulture,
									"Exception at: {0} index: {1}",
									path,
									index.ToString(
										CultureInfo.InvariantCulture));

								Log.Error(message);
								Log.Error(exception.ToString());
							}
						}

						processed += folderAction(folder, condition);
					}
				}
				catch (COMException exception)
				{
					string path = GetFolderPath(folder);

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"Exception at: {0}",
						path);

					Log.Error(message);
					Log.Error(exception.ToString());
				}
			}

			return processed;
		}

		/// <summary>
		/// Recurse folders.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <param name="condition">A conditional to check.</param>
		/// <param name="folderAction">The delegate to act upon.</param>
		/// <returns>A value processed from the delegate.</returns>
		public static async Task<int> RecurseFoldersAsync(
			MAPIFolder folder,
			bool condition,
			FolderActionAsync folderAction)
		{
			int processed = 0;

			if (folder != null && folderAction != null)
			{
				try
				{
					bool isDeletedFolder = IsDeletedFolder(folder);

					// Skip processing of system deleted items folder.
					if (isDeletedFolder == false)
					{
						int folderCount = folder.Folders.Count;

						// Office uses 1 based indexes from VBA.
						// Iterate in reverse order as the group may change.
						for (int index = folderCount; index > 0; index--)
						{
							try
							{
								MAPIFolder subFolder = folder.Folders[index];

								processed += await RecurseFoldersAsync(
									subFolder,
									condition,
									folderAction).ConfigureAwait(false);

								Marshal.ReleaseComObject(subFolder);
							}
							catch (COMException exception)
							{
								string path = GetFolderPath(folder);

								string message = string.Format(
									CultureInfo.InvariantCulture,
									"Exception at: {0} index: {1}",
									path,
									index.ToString(CultureInfo.InvariantCulture));

								Log.Error(message);
								Log.Error(exception.ToString());
							}
						}

						processed += await folderAction(folder, condition).
							ConfigureAwait(false);
					}
				}
				catch (COMException exception)
				{
					string path = GetFolderPath(folder);

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"Exception at: {0}",
						path);

					Log.Error(message);
					Log.Error(exception.ToString());
				}
			}

			return processed;
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="folder">The current folder.</param>
		/// <param name="condition">A condition to check for. Currently
		/// unused. Set here to match delegate signature.</param>
		/// <returns>The count of removed folders.</returns>
		public static int RemoveEmptyFolders(
			MAPIFolder folder, bool condition)
		{
			int removedFolders = 0;

			if (folder != null)
			{
				removedFolders =
					RecurseFolders(folder, condition, RemoveEmptyFolder);
			}

			return removedFolders;
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="folder">The current folder.</param>
		/// <param name="condition">A condition to check for. Currently
		/// unused. Set here to match delegate signature.</param>
		/// <returns>The count of removed folders.</returns>
		public static async Task<int> RemoveEmptyFoldersAsync(
			MAPIFolder folder, bool condition)
		{
			int removedFolders = 0;

			if (folder != null)
			{
				removedFolders = await
					RecurseFoldersAsync(
						folder, condition, RemoveEmptyFolderAsync).
						ConfigureAwait(false);
			}

			return removedFolders;
		}

		/// <summary>
		/// Remove folder from PST store.
		/// </summary>
		/// <param name="folder">The folder to remove.</param>
		/// <param name="subFolderIndex">The index of the sub-folder.</param>
		/// <param name="force">Whether to force the removal.</param>
		public static void RemoveFolder(
			MAPIFolder folder,
			int subFolderIndex,
			bool force)
		{
			if (folder != null)
			{
				string path = GetFolderPath(folder);

				bool isReserved = IsReservedFolder(folder);

				if (isReserved == false)
				{
					// Perhaps because interaction through COM interop, the count
					// values sometimes seem a bit behind, so pause a little bit
					// before moving on.
					System.Threading.Thread.Sleep(400);

					bool empty = true;
					string message = "Removing empty folder: " + path;

					if (folder.Folders.Count > 0 || folder.Items.Count > 0)
					{
						Log.Warn(
							"Attempting to remove non empty folder: " + path);
						empty = false;
					}

					if (force == true)
					{
						message = "Forcing removal of folder: " + path;
					}

					if (force == true || empty == true)
					{
						Log.Info(message);

						MAPIFolder parentFolder = GetParent(folder);

						if (parentFolder != null)
						{
							parentFolder.Folders.Remove(subFolderIndex);
						}
					}
				}
			}
		}

		/// <summary>
		/// Remove folder from PST store.
		/// </summary>
		/// <param name="path">The path of current folder.</param>
		/// <param name="subFolderIndex">The index of the sub-folder.</param>
		/// <param name="subFolder">The sub-folder.</param>
		/// <param name="force">Whether to force the removal.</param>
		[Obsolete("RemoveFolder(string, int, MAPIFolder, bool) is deprecated," +
			" please use RemoveFolder(MAPIFolder, int, bool) instead.")]
		public static void RemoveFolder(
			string path,
			int subFolderIndex,
			MAPIFolder subFolder,
			bool force)
		{
			RemoveFolder(subFolder, subFolderIndex, force);
		}

		/// <summary>
		/// Safely delete the folder.
		/// </summary>
		/// <param name="folder">The folder to delete.</param>
		/// <returns>Indicates whether the folder was actually deleted
		/// or not.</returns>
		public static bool SafeDelete(MAPIFolder folder)
		{
			bool isDeleted = false;

			if (folder != null)
			{
				if (folder.Folders.Count == 0 && folder.Items.Count == 0)
				{
					bool isReservedFolder = IsReservedFolder(folder);

					if (isReservedFolder == true)
					{
						string name = folder.Name;
						Log.Warn("Not deleting reserved folder: " +
							name);
					}
					else
					{
						folder.Delete();
						isDeleted = true;
					}
				}
			}

			return isDeleted;
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
					try
					{
						NameSpace session = outlookAccount.Session;

						MailItem item = session.OpenSharedItem(filePath);

						item.UnRead = false;
						item.Save();

						item.Move(pstFolder);

						Marshal.ReleaseComObject(item);
					}
					catch (COMException exception)
					{
						Log.Error(exception.ToString());
					}
				}
				else
				{
					Log.Warn("File doesn't exist: " + filePath);
				}
			}
		}

		/// <summary>
		/// Get a list of item hashes from the given folder.
		/// </summary>
		/// <param name="folder">The MAPI folder to process.</param>
		/// <returns>A list of item hashes from the given folder.</returns>
		public IDictionary<string, IList<string>> GetItemHashes(
			MAPIFolder folder)
		{
			if (folder != null)
			{
				bool isDeletedFolder = IsDeletedFolder(folder);

				// Skip processing of system deleted items folder.
				if (isDeletedFolder == false)
				{
					int folderCount = folder.Folders.Count;

					// Office uses 1 based indexes from VBA.
					// Iterate in reverse order as the group may change.
					for (int index = folderCount; index > 0; index--)
					{
						MAPIFolder subFolder = folder.Folders[index];

						storeHashTable = GetItemHashes(subFolder);

						Marshal.ReleaseComObject(subFolder);
					}

					storeHashTable = GetFolderHashTable(folder);
				}
			}

			return storeHashTable;
		}

		/// <summary>
		/// Get senders counts.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>The count of each sender.</returns>
		public IDictionary<string, int> GetSendersCount(
			MAPIFolder folder)
		{
			if (folder != null && sendersCounts != null)
			{
				Folders folders = folder.Folders;
				int count = folders.Count;

				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group may change.
				for (int index = count; index > 0; index--)
				{
					MAPIFolder subFolder = folder.Folders[index];

					sendersCounts = GetSendersCount(subFolder);

					Marshal.ReleaseComObject(subFolder);
				}

				string path = GetFolderPath(folder);

				Items items = folder.Items;
				int total = items.Count;
				string totals = total.ToString(CultureInfo.InvariantCulture);
				Log.Info("Checking senders in: " + path + ": " + totals);

				sendersCounts = GetFolderSendersCount(folder);
			}

			return sendersCounts;
		}

		/// <summary>
		/// Merge folders.
		/// </summary>
		/// <param name="folder">The current folder.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		public void MergeFolders(MAPIFolder folder, bool dryRun)
		{
			if (folder != null)
			{
				RecurseFolders(folder, dryRun, MergeThisFolder);
			}
		}

		/// <summary>
		/// Merge folders.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="folder">The current folder.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		[Obsolete("MergeFolders(string, MAPIFolder, bool) is deprecated, " +
			"please use MergeFolders(MAPIFolder, bool) instead.")]
		public void MergeFolders(string path, MAPIFolder folder, bool dryRun)
		{
			if (folder != null)
			{
				RecurseFolders(folder, dryRun, MergeThisFolder);
			}
		}

		/// <summary>
		/// Merge folders.
		/// </summary>
		/// <param name="folder">The current folder.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MergeFoldersAsync(MAPIFolder folder, bool dryRun)
		{
			if (folder != null)
			{
				await RecurseFoldersAsync(
					folder, dryRun, MergeThisFolderAsync).
					ConfigureAwait(false);
			}
		}

		/// <summary>
		/// Move the folder contents.
		/// </summary>
		/// <param name="source">The source folder.</param>
		/// <param name="destination">The destination folder.</param>
		public void MoveFolderContents(
			MAPIFolder source, MAPIFolder destination)
		{
			if (source != null && destination != null)
			{
				string destinationPath = GetFolderPath(destination);

				string sourceName = source.Name;
				string destinationName = destination.Name;

				LogFormatMessage.Info(
					"{0}: Merging {1} into {2}",
					destinationPath,
					sourceName,
					destinationName);

				MoveFolderItems(source, destination);
				MoveSubFolders(source, destination);
			}
		}

		/// <summary>
		/// Move the folder contents.
		/// </summary>
		/// <param name="path">Path of parent folder.</param>
		/// <param name="source">The source folder.</param>
		/// <param name="destination">The destination folder.</param>
		[Obsolete("MoveFolderContents(string, MAPIFolder, MAPIFolder) " +
			"is deprecated, please use " +
			"MoveFolderContents(MAPIFolder, MAPIFolder) instead.")]
		public void MoveFolderContents(
			string path, MAPIFolder source, MAPIFolder destination)
		{
			MoveFolderContents(source, destination);
		}

		/// <summary>
		/// Move the folder contents.
		/// </summary>
		/// <param name="source">The source folder.</param>
		/// <param name="destination">The destination folder.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MoveFolderContentsAsync(
			MAPIFolder source, MAPIFolder destination)
		{
			if (source != null && destination != null)
			{
				string destinationPath = GetFolderPath(destination);

				string sourceName = source.Name;
				string destinationName = destination.Name;

				LogFormatMessage.Info(
					"{0}: Merging {1} into {2}",
					destinationPath,
					sourceName,
					destinationName);

				await MoveFolderItemsAsync(source, destination).
					ConfigureAwait(false);
				await MoveSubFoldersAsync(source, destination).
					ConfigureAwait(false);
			}
		}

		/// <summary>
		/// Remove duplicates items from the given folder.
		/// </summary>
		/// <param name="folder">The MAPI folder to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>The total count of duplicates removed.</returns>
		public int RemoveDuplicates(MAPIFolder folder, bool dryRun)
		{
			if (folder != null)
			{
				RemovedDuplicates = RecurseFolders(
					folder, dryRun, RemoveDuplicatesFromThisFolder);
			}

			return RemovedDuplicates;
		}

		/// <summary>
		/// Remove duplicates items from the given folder.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="folder">The MAPI folder to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>The total count of duplicates removed.</returns>
		[Obsolete("RemoveDuplicates(string, MAPIFolder, bool) is deprecated," +
			" please use RemoveDuplicates(MAPIFolder, bool) instead.")]
		public int RemoveDuplicates(
			string path, MAPIFolder folder, bool dryRun)
		{
			return RemoveDuplicates(folder, dryRun);
		}

		/// <summary>
		/// Remove duplicates items from the given folder.
		/// </summary>
		/// <param name="folder">The MAPI folder to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>The total count of duplicates removed.</returns>
		public async Task<int> RemoveDuplicatesAsync(
			MAPIFolder folder, bool dryRun)
		{
			if (folder != null)
			{
				RemovedDuplicates = await RecurseFoldersAsync(
					folder, dryRun, RemoveDuplicatesFromThisFolderAsync).
						ConfigureAwait(false);
			}

			return RemovedDuplicates;
		}

		private static IDictionary<string, IList<string>> AddHashToTable(
			IDictionary<string, IList<string>> hashTable,
			string hash,
			string entryId)
		{
			if (!string.IsNullOrEmpty(hash))
			{
				bool keyExists = hashTable.ContainsKey(hash);

				if (keyExists == true)
				{
					IList<string> bucket = hashTable[hash];
					bucket.Add(entryId);
				}
				else
				{
					IList<string> bucket = new List<string>();
					bucket.Add(entryId);

					hashTable.Add(hash, bucket);
				}
			}

			return hashTable;
		}

		private static string CheckFolderNameNormalization(string folderName)
		{
			string duplicatePattern = null;
			string[] duplicatePatterns =
			{
				@"\s*\(\d*?\)$", @"^\s+(?=[a-zA-Z])+", @"^_+(?=[a-zA-Z])+",
				@"_\d$", @"(?<=[a-zA-Z0-9])_$", @"^[a-fA-F]{1}\d{1}_",

				// Matches Something  ab2 (2 spaces and 2 or 3 hex numbers)
				@"(?<=[a-zA-Z0-9&,])\s{2,3}[0-9a-fA-F]{2,3}$",

				// Matches Something ab2 (at least 1 space and 3 hex numbers)
				@"(?<=[a-zA-Z0-9&-,])\s+[0-9a-fA-F]{3}$",

				// Matches Something@ 896
				// (at least 1 space and 2 or 3 hex numbers)
				@"(?<=[a-zA-Z0-9&,])@\s+[0-9a-fA-F]{2,3}$",

				// Matches Something - 77f (1 space and 2 or3 hex numbers)
				@"(?<=[a-zA-Z0-9&,])\s{1}-\s{1}[0-9a-fA-F]{2,3}$",
				@"\s*-\s*Copy$", @"^[A-F]{1}_"
			};

			foreach (string pattern in duplicatePatterns)
			{
				if (Regex.IsMatch(folderName, pattern))
				{
					duplicatePattern = pattern;
					break;
				}
			}

			return duplicatePattern;
		}

		private static bool DoesSiblingFolderExist(
			MAPIFolder folder, string folderName)
		{
			bool folderExists = false;
			MAPIFolder parentFolder = GetParent(folder);

			if (parentFolder != null)
			{
				folderExists = DoesFolderExist(parentFolder, folderName);

				Marshal.ReleaseComObject(parentFolder);
			}

			return folderExists;
		}

		private static string GetNormalizedFolderName(
			string folderName, string pattern)
		{
			string newFolderName = Regex.Replace(
				folderName,
				pattern,
				string.Empty,
				RegexOptions.ExplicitCapture);

			return newFolderName;
		}

		private static MAPIFolder GetPathFolder(
			Store store, string path, bool justParent)
		{
			MAPIFolder currentFolder = null;

			if (store != null)
			{
				currentFolder = store.GetRootFolder();

				// If no folder path given, start with the root folder.
				if (!string.IsNullOrWhiteSpace(path))
				{
					string[] parts = GetPathParts(path);

					int maxParts = parts.Length;

					if (justParent == true)
					{
						maxParts = parts.Length - 1;
					}

					for (int index = 0; index < maxParts; index++)
					{
						string part = parts[index];

						if (index == 0)
						{
							string rootFolderName = currentFolder.Name;

							if (part.Equals(
								rootFolderName,
								StringComparison.OrdinalIgnoreCase))
							{
								// root, so skip over
								continue;
							}
						}

						currentFolder = AddFolder(currentFolder, part);
					}
				}
			}

			return currentFolder;
		}

		private static string[] GetPathParts(string path)
		{
			path = RemoveStoreFromPath(path);

			char[] charSeparators = new char[] { '\\', '/' };
			string[] parts = path.Split(
				charSeparators, StringSplitOptions.RemoveEmptyEntries);

			return parts;
		}

		private static void ItemsIterator(
			MAPIFolder source,
			MAPIFolder destination,
			ItemAction itemAction,
			string messageTemplate)
		{
			Items items = source.Items;

			int ascendingCount = 1;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = items.Count; index > 0; index--)
			{
				object item = items[index];

				int sectionIndicator = ascendingCount % 100;

				if (ascendingCount == 1 || sectionIndicator == 0)
				{
					Log.Info(
						messageTemplate +
						ascendingCount.ToString(CultureInfo.InvariantCulture));
				}

				itemAction(destination, item);

				ascendingCount++;
			}
		}

		private static async Task ItemsIteratorAsync(
			MAPIFolder source,
			MAPIFolder destination,
			ItemActionAsync itemAction,
			string messageTemplate)
		{
			Items items = source.Items;

			int ascendingCount = 1;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = items.Count; index > 0; index--)
			{
				object item = items[index];

				int sectionIndicator = ascendingCount % 100;

				if (ascendingCount == 1 || sectionIndicator == 0)
				{
					Log.Info(
						messageTemplate +
						ascendingCount.ToString(CultureInfo.InvariantCulture));
				}

				await itemAction(destination, item).ConfigureAwait(false);

				ascendingCount++;
			}
		}

		private static void ListItem(MailItem mailItem, string prefixMessage)
		{
			string sentOn = mailItem.SentOn.ToString(
				"yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

			LogFormatMessage.Info(
				"{0} {1}: From: {2}: {3} Subject: {4}",
				prefixMessage,
				sentOn,
				mailItem.SenderName,
				mailItem.SenderEmailAddress,
				mailItem.Subject);
		}

		private static bool MergeDeletedItemsFolder(MAPIFolder folder)
		{
			bool removed = false;
			string name = folder.Name;

			if (DeletedFolders.Contains(name))
			{
				removed = SafeDelete(folder);
			}

			return removed;
		}

		private static void MoveItem(MAPIFolder destination, object item)
		{
			MapiItem.Moveitem(item, destination);
		}

		private static async Task MoveItemAsync(
			MAPIFolder destination, object item)
		{
			await MapiItem.MoveitemAsync(item, destination).
				ConfigureAwait(false);
		}

		private static void MoveFolderItems(
			MAPIFolder source, MAPIFolder destination)
		{
			ItemsIterator(
				source, destination, MoveItem, "Moving Items from: ");
		}

		private static async Task MoveFolderItemsAsync(
			MAPIFolder source, MAPIFolder destination)
		{
			await ItemsIteratorAsync(
				source, destination, MoveItemAsync, "Moving Items from: ").
				ConfigureAwait(false);
		}

		private static int RemoveEmptyFolder(MAPIFolder folder, bool condition)
		{
			int count = 0;
			bool isDeleted = SafeDelete(folder);

			if (isDeleted == true)
			{
				count = 1;
			}

			return count;
		}

		private static Task<int> RemoveEmptyFolderAsync(
			MAPIFolder folder, bool condition)
		{
			int count = 0;
			bool isDeleted = SafeDelete(folder);

			if (isDeleted == true)
			{
				count = 1;
			}

			return Task.FromResult(count);
		}

		private static string RemoveStoreFromPath(string path)
		{
			if (path.Contains("::", StringComparison.OrdinalIgnoreCase))
			{
				int position = path.IndexOf(
					"::", StringComparison.OrdinalIgnoreCase);
				position += 2;

				path = path[position..];
			}

			return path;
		}

		private void AddItemHashToTable(MAPIFolder folder, object item)
		{
			switch (item)
			{
				// Initially, just focus on MailItems
				case MailItem mailItem:
					string entryId = mailItem.EntryID;
					string hash = MapiItem.GetItemHash(mailItem);

					storeHashTable =
						AddHashToTable(storeHashTable, hash, entryId);

					Marshal.ReleaseComObject(mailItem);
					break;
				default:
					Log.Info("Ignoring item of non-MailItem type: ");
					break;
			}
		}

		private async Task AddItemHashToTableAsync(
			MAPIFolder folder, object item)
		{
			switch (item)
			{
				// Initially, just focus on MailItems
				case MailItem mailItem:
					string entryId = mailItem.EntryID;
					string hash = await MapiItem.GetItemHashAsync(mailItem).
						ConfigureAwait(false);

					storeHashTable =
						AddHashToTable(storeHashTable, hash, entryId);

					Marshal.ReleaseComObject(mailItem);
					break;
				default:
					Log.Info("Ignoring item of non-MailItem type: ");
					break;
			}
		}

		private void AddSenderCount(MAPIFolder folder, object item)
		{
			switch (item)
			{
				case MailItem mailItem:
					string sender = mailItem.SenderEmailAddress;

					if (!string.IsNullOrWhiteSpace(sender))
					{
						if (sendersCounts.ContainsKey(sender))
						{
							sendersCounts[sender]++;
						}
						else
						{
							sendersCounts.Add(sender, 1);
						}
					}
					else
					{
						string subject = mailItem.Subject;
						Log.Warn(
							"Item has no sender - subject:" + subject);
					}

					Marshal.ReleaseComObject(mailItem);
					break;
				default:
					Log.Info("Ignoring item of non-MailItem type: ");
					break;
			}
		}

		private void CheckForDuplicateFolders(MAPIFolder folder, bool dryRun)
		{
			string folderName = folder.Name;

			string duplicatePattern = CheckFolderNameNormalization(folderName);

			if (!string.IsNullOrWhiteSpace(duplicatePattern))
			{
				MergeDuplicateFolder(folder, duplicatePattern, dryRun);
			}
		}

		private async Task CheckForDuplicateFoldersAsync(
			MAPIFolder folder, bool dryRun)
		{
			string folderName = folder.Name;

			string duplicatePattern = CheckFolderNameNormalization(folderName);

			if (!string.IsNullOrWhiteSpace(duplicatePattern))
			{
				await MergeDuplicateFolderAsync(
					folder, duplicatePattern, dryRun).ConfigureAwait(false);
			}
		}

		private int DeleteDuplicates(IList<string> duplicateSet, bool dryRun)
		{
			string keeper = duplicateSet[0];
			duplicateSet.RemoveAt(0);

			// Count only the ones to remove.
			int removeDuplicates = duplicateSet.Count;

			NameSpace session = outlookAccount.Session;

			MailItem mailItem = session.GetItemFromID(keeper);
			string keeperSynopses = MapiItem.GetItemSynopses(mailItem);

			string message = string.Format(
				CultureInfo.InvariantCulture,
				"{0} Duplicates Found for: ",
				removeDuplicates.ToString(CultureInfo.InvariantCulture));

			ListItem(mailItem, message);

			foreach (string duplicateId in duplicateSet)
			{
				MapiItem.DeleteDuplicate(
					session, duplicateId, keeperSynopses, dryRun);
			}

			return removeDuplicates;
		}

		private IDictionary<string, IList<string>> GetFolderHashTable(
			MAPIFolder folder)
		{
			IDictionary<string, IList<string>> hashTable = null;

			if (folder != null)
			{
				hashTable = new Dictionary<string, IList<string>>();

				ItemsIterator(
					folder,
					folder,
					AddItemHashToTable,
					"Getting Item Hashes from: ");
			}

			return storeHashTable;
		}

		private async Task<IDictionary<string, IList<string>>>
			GetFolderHashTableAsync(MAPIFolder folder)
		{
			IDictionary<string, IList<string>> hashTable = null;

			if (folder != null)
			{
				hashTable = new Dictionary<string, IList<string>>();

				await ItemsIteratorAsync(
					folder,
					folder,
					AddItemHashToTableAsync,
					"Getting Item Hashes from: ").ConfigureAwait(false);
			}

			return storeHashTable;
		}

		private IDictionary<string, int> GetFolderSendersCount(
			MAPIFolder folder)
		{
			if (folder != null)
			{
				ItemsIterator(
					folder,
					folder,
					AddSenderCount,
					"Getting Item Senders Count from: ");
			}

			return sendersCounts;
		}

		private void MergeDuplicateFolder(
			MAPIFolder folder,
			string duplicatePattern,
			bool dryRun)
		{
			string newFolderName =
				GetNormalizedFolderName(folder.Name, duplicatePattern);

			bool folderExists = DoesSiblingFolderExist(folder, newFolderName);

			string source = folder.Name;

			if (folderExists == true)
			{
				if (dryRun == true)
				{
					Log.Info(
						"WOULD merge " + source + " into " + newFolderName);
				}
				else
				{
					MAPIFolder parentFolder = GetParent(folder);

					if (parentFolder != null)
					{
						// Move items
						MAPIFolder destination =
							parentFolder.Folders[newFolderName];

						MoveFolderContents(folder, destination);

						// Once all the items have been moved,
						// remove the folder.
						SafeDelete(folder);
					}
				}
			}
			else
			{
				if (dryRun == true)
				{
					Log.Info("WOULD move " + source + " to " + newFolderName);
				}
				else
				{
					try
					{
						Log.Info("Moving " + source + " to " + newFolderName);
						folder.Name = newFolderName;
					}
					catch (COMException)
					{
						string message = string.Format(
							CultureInfo.InvariantCulture,
							"Failed renaming {0} to {1} with COMException",
							folder.Name,
							newFolderName);
						Log.Error(message);
					}
				}
			}
		}

		private async Task MergeDuplicateFolderAsync(
			MAPIFolder folder,
			string duplicatePattern,
			bool dryRun)
		{
			string newFolderName =
				GetNormalizedFolderName(folder.Name, duplicatePattern);

			bool folderExists = DoesSiblingFolderExist(folder, newFolderName);

			string source = folder.Name;

			if (folderExists == true)
			{
				if (dryRun == true)
				{
					Log.Info(
						"WOULD merge " + source + " into " + newFolderName);
				}
				else
				{
					MAPIFolder parentFolder = GetParent(folder);

					if (parentFolder != null)
					{
						// Move items
						MAPIFolder destination =
							parentFolder.Folders[newFolderName];

						await MoveFolderContentsAsync(folder, destination).
							ConfigureAwait(false);

						// Once all the items have been moved,
						// remove the folder.
						SafeDelete(folder);
					}
				}
			}
			else
			{
				if (dryRun == true)
				{
					Log.Info("WOULD move " + source + " to " + newFolderName);
				}
				else
				{
					try
					{
						Log.Info("Moving " + source + " to " + newFolderName);
						folder.Name = newFolderName;
					}
					catch (COMException)
					{
						string message = string.Format(
							CultureInfo.InvariantCulture,
							"Failed renaming {0} to {1} with COMException",
							folder.Name,
							newFolderName);
						Log.Error(message);
					}
				}
			}
		}

		private void MergeFolderWithParent(MAPIFolder folder, bool dryRun)
		{
			string name = folder.Name;
			MAPIFolder parent = GetParent(folder);

			if (parent != null)
			{
				string path = GetFolderPath(folder);

				if (dryRun == true)
				{
					Log.Info(
						"At: " + path + " WOULD Move into parent: " + name);
				}
				else
				{
					Log.Info("At: " + path + " Moving into parent: " + name);

					MoveFolderContents(folder, parent);

					// Once all the items have been moved,
					// now remove the folder.
					SafeDelete(folder);
				}
			}
		}

		private async Task MergeFolderWithParentAsync(
			MAPIFolder folder, bool dryRun)
		{
			string name = folder.Name;
			MAPIFolder parent = GetParent(folder);

			if (parent != null)
			{
				string path = GetFolderPath(folder);

				if (dryRun == true)
				{
					Log.Info(
						"At: " + path + " WOULD Move into parent: " + name);
				}
				else
				{
					Log.Info("At: " + path + " Moving into parent: " + name);

					await MoveFolderContentsAsync(folder, parent).
						ConfigureAwait(false);

					// Once all the items have been moved,
					// now remove the folder.
					SafeDelete(folder);
				}
			}
		}

		private int MergeThisFolder(MAPIFolder folder, bool dryRun)
		{
			int processed = 0;
			CheckForDuplicateFolders(folder, dryRun);

			bool removed = MergeDeletedItemsFolder(folder);

			if (removed == false)
			{
				bool topLevel = IsTopLevelFolder(folder);

				if (topLevel == false)
				{
					string name = folder.Name;
					MAPIFolder parent = GetParent(folder);

					if (parent != null)
					{
						string parentName = parent.Name;

						if (parentName.Equals(
							name, StringComparison.OrdinalIgnoreCase))
						{
							MergeFolderWithParent(folder, dryRun);
							processed = 1;
						}
					}
				}
			}
			else
			{
				processed = 1;
			}

			return processed;
		}

		private async Task<int> MergeThisFolderAsync(
			MAPIFolder folder, bool dryRun)
		{
			int processed = 0;
			await CheckForDuplicateFoldersAsync(folder, dryRun).
				ConfigureAwait(false);

			bool removed = MergeDeletedItemsFolder(folder);

			if (removed == false)
			{
				bool topLevel = IsTopLevelFolder(folder);

				if (topLevel == false)
				{
					string name = folder.Name;
					MAPIFolder parent = GetParent(folder);

					if (parent != null)
					{
						string parentName = parent.Name;

						if (parentName.Equals(
							name, StringComparison.OrdinalIgnoreCase))
						{
							await MergeFolderWithParentAsync(
								folder, dryRun).ConfigureAwait(false);
							processed = 1;
						}
					}
				}
			}
			else
			{
				processed = 1;
			}

			return processed;
		}

		private void MoveSubFolders(MAPIFolder source, MAPIFolder destination)
		{
			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = source.Folders.Count; index > 0; index--)
			{
				MAPIFolder subFolder = source.Folders[index];

				MoveFolder(subFolder, destination, index);
			}
		}

		private async Task MoveSubFoldersAsync(
			MAPIFolder source, MAPIFolder destination)
		{
			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = source.Folders.Count; index > 0; index--)
			{
				MAPIFolder subFolder = source.Folders[index];

				await MoveFolderAsync(subFolder, destination, index).
					ConfigureAwait(false);
			}
		}

		private void MoveFolder(
			MAPIFolder source, MAPIFolder destinationParent, int index)
		{
			string destinationParentPath = GetFolderPath(destinationParent);

			string destinationName = destinationParent.Name;

			string name = source.Name;
			MAPIFolder destinationSubFolder =
				GetSubFolder(destinationParent, name);

			if (destinationSubFolder == null)
			{
				// Folder doesn't already exist, so just move it.
				LogFormatMessage.Info(
					"at: {0} Moving {1} to {2}",
					destinationParentPath,
					name,
					destinationName);

				try
				{
					// In some rare occasions, the folder is actually already
					// deleted, but isn't acknowledged in time, but by the
					// time the process gets to here, it seems deleted. Thus,
					// trying to move the folder is going to cause an
					// exception.  Just catch it and move on.
					source.MoveTo(destinationParent);
				}
				catch (COMException exception)
				{
					Log.Warn(exception.ToString());
				}
			}
			else
			{
				// Folder exists, so if just moving it, it will get
				// renamed something FolderName (2), so need to merge.
				string subPath = destinationParentPath + "/" + source.Name;

				LogFormatMessage.Info(
					"at: {0} Merging {1} to {2}",
					subPath,
					name,
					destinationName);

				MoveFolderContents(source, destinationSubFolder);

				// Once all the items have been moved,
				// now remove the folder.
				RemoveFolder(source, index, false);
			}
		}

		private async Task MoveFolderAsync(
			MAPIFolder source, MAPIFolder destinationParent, int index)
		{
			string destinationParentPath = GetFolderPath(destinationParent);

			string destinationName = destinationParent.Name;

			string name = source.Name;
			MAPIFolder destinationSubFolder =
				GetSubFolder(destinationParent, name);

			if (destinationSubFolder == null)
			{
				// Folder doesn't already exist, so just move it.
				LogFormatMessage.Info(
					"at: {0} Moving {1} to {2}",
					destinationParentPath,
					name,
					destinationName);

				try
				{
					// In some rare occasions, the folder is actually already
					// deleted, but isn't acknowledged in time, but by the
					// time the process gets to here, it seems deleted. Thus,
					// trying to move the folder is going to cause an
					// exception.  Just catch it and move on.
					source.MoveTo(destinationParent);
				}
				catch (COMException exception)
				{
					Log.Warn(exception.ToString());
				}
			}
			else
			{
				// Folder exists, so if just moving it, it will get
				// renamed something FolderName (2), so need to merge.
				string subPath = destinationParentPath + "/" + source.Name;

				LogFormatMessage.Info(
					"at: {0} Merging {1} to {2}",
					subPath,
					name,
					destinationName);

				await MoveFolderContentsAsync(source, destinationSubFolder).
					ConfigureAwait(false);

				// Once all the items have been moved,
				// now remove the folder.
				RemoveFolder(source, index, false);
			}
		}

		private int RemoveDuplicatesFromThisFolder(
			MAPIFolder folder, bool dryRun)
		{
			int removedDuplicates = 0;

			IDictionary<string, IList<string>> hashTable =
				GetFolderHashTable(folder);

			var duplicates = hashTable.Where(p => p.Value.Count > 1);
			int duplicateCount = duplicates.Count();

			if (duplicateCount > 0)
			{
				string path = GetFolderPath(folder);
				Log.Info("Duplicates found at: " + path);
			}

			foreach (KeyValuePair<string, IList<string>> duplicateSet in
				duplicates)
			{
				removedDuplicates +=
					DeleteDuplicates(duplicateSet.Value, dryRun);
			}

			Marshal.ReleaseComObject(folder);

			return removedDuplicates;
		}

		private async Task<int> RemoveDuplicatesFromThisFolderAsync(
			MAPIFolder folder, bool dryRun)
		{
			int duplicateCount = 0;

			string path = GetFolderPath(folder);

			IDictionary<string, IList<string>> hashTable =
				await GetFolderHashTableAsync(folder).ConfigureAwait(false);

			IEnumerable<KeyValuePair<string, IList<string>>> duplicates =
				hashTable.Where(p => p.Value.Count > 1);
			int duplicateSetCount = duplicates.Count();

			if (duplicateSetCount > 0)
			{
				Log.Info("Duplicates found at: " + path);
			}

			foreach (KeyValuePair<string, IList<string>> duplicateSet in
				duplicates)
			{
				duplicateCount +=
					DeleteDuplicates(duplicateSet.Value, dryRun);
			}

			Marshal.ReleaseComObject(folder);

			return duplicateCount;
		}
	}
}
