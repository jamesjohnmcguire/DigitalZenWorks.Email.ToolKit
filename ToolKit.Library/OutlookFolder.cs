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

namespace DigitalZenWorks.Email.ToolKit
{
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

		private uint removedFolders;
		private uint totalFolders;

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
					Log.Info("Adding outlook folder: " + folderName);

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
		/// Get the folder's full path.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>The folder's full path.</returns>
		public static string GetFolderPath(MAPIFolder folder)
		{
			string path = null;

			if (folder != null)
			{
				path = folder.Name;

				while (folder.Parent is not null &&
					folder.Parent is MAPIFolder)
				{
					folder = folder.Parent;
					string name = folder.Name;
					path = name + "/" + path;
				}

				string storeName = OutlookStore.GetStoreName(folder.Store);
				path = storeName + "::" + path;
			}

			return path;
		}

		/// <summary>
		/// Get sub folder from parent.
		/// </summary>
		/// <param name="parentFolder">The parent folder.</param>
		/// <param name="folderName">The new folder name.</param>
		/// <returns>The added folder.</returns>
		public static MAPIFolder GetSubFolder(
			MAPIFolder parentFolder, string folderName)
		{
			MAPIFolder pstFolder = null;

			if (parentFolder != null && !string.IsNullOrWhiteSpace(folderName))
			{
				try
				{
					pstFolder = parentFolder.Folders[folderName];
				}
				catch (COMException)
				{
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
					if (folder.Parent is not null &&
						folder.Parent is MAPIFolder)
					{
						MAPIFolder parent = folder.Parent;

						// Check if root folder
						if (parent.Parent is null ||
							parent.Parent is not MAPIFolder)
						{
							isDeletedFolder = true;
						}
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
					if (folder.Parent is not null &&
						folder.Parent is MAPIFolder)
					{
						MAPIFolder parent = folder.Parent;

						// Check if root folder
						if (parent.Parent is null ||
							parent.Parent is not MAPIFolder)
						{
							reserved = true;
						}
					}
				}
				else if (folder.Parent is null ||
					folder.Parent is not MAPIFolder)
				{
					// root folder
					reserved = true;
				}
			}

			return reserved;
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
		/// Merge folders.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="folder">The current folder.</param>
		public void MergeFolders(string path, MAPIFolder folder)
		{
			if (folder != null)
			{
				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group may change.
				for (int index = folder.Folders.Count; index > 0; index--)
				{
					MAPIFolder subFolder = folder.Folders[index];

					string subPath = path + "/" + subFolder.Name;

					MergeFolders(subPath, subFolder);

					CheckForDuplicateFolders(path, index, subFolder);

					totalFolders++;
					Marshal.ReleaseComObject(subFolder);
				}
			}
		}

		/// <summary>
		/// Move the folder contents.
		/// </summary>
		/// <param name="path">Path of parent folder.</param>
		/// <param name="source">The source folder.</param>
		/// <param name="destination">The destination folder.</param>
		public void MoveFolderContents(
			string path, MAPIFolder source, MAPIFolder destination)
		{
			if (source != null && destination != null)
			{
				string sourceName = source.Name;
				string destinationName = destination.Name;

				string message = string.Format(
					CultureInfo.InvariantCulture,
					"{0}: Merging {1} into {2}",
					path,
					sourceName,
					destinationName);
				Log.Info(message);

				MoveFolderItems(source, destination);
				MoveSubFolders(path, source, destination);
			}
		}

		/// <summary>
		/// Remove duplicates items from the given folder.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="folder">The MAPI folder to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>An array of duplicate sets and total duplicate items
		/// count.</returns>
		public int[] RemoveDuplicates(
			string path, MAPIFolder folder, bool dryRun)
		{
			int[] duplicateCounts = new int[2];

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

						int[] subFolderduplicateCounts =
							RemoveDuplicates(path, subFolder, dryRun);

						duplicateCounts[0] += subFolderduplicateCounts[0];
						duplicateCounts[1] += subFolderduplicateCounts[1];

						totalFolders++;
						Marshal.ReleaseComObject(subFolder);
					}

					int[] duplicateCountsThisFolder =
						RemoveDuplicatesFromThisFolder(folder, dryRun);

					duplicateCounts[0] += duplicateCountsThisFolder[0];
					duplicateCounts[1] += duplicateCountsThisFolder[1];
				}
			}

			return duplicateCounts;
		}

		/// <summary>
		/// Remove folder from PST store.
		/// </summary>
		/// <param name="path">The path of current folder.</param>
		/// <param name="subFolderIndex">The index of the sub-folder.</param>
		/// <param name="subFolder">The sub-folder.</param>
		/// <param name="force">Whether to force the removal.</param>
		public void RemoveFolder(
			string path,
			int subFolderIndex,
			MAPIFolder subFolder,
			bool force)
		{
			if (subFolder != null)
			{
				// Perhaps because interaction through COM interop, the count
				// values sometimes seem a bit behind, so pause a little bit
				// before moving on.
				System.Threading.Thread.Sleep(100);

				if (subFolder.Folders.Count > 0 || subFolder.Items.Count > 0)
				{
					Log.Warn("Attempting to remove non empty folder: " + path);
				}

				if (force == true || (subFolder.Folders.Count == 0 &&
					subFolder.Items.Count == 0))
				{
					Log.Info("Removing empty folder: " + path);

					bool isReserved = IsReservedFolder(subFolder);

					if (isReserved == false)
					{
						MAPIFolder parentFolder = subFolder.Parent;

						parentFolder.Folders.Remove(subFolderIndex);
					}

					removedFolders++;
				}
			}
		}

		private static bool DoesSiblingFolderExist(
			MAPIFolder folder, string folderName)
		{
			MAPIFolder parentFolder = folder.Parent;

			bool folderExists = DoesFolderExist(parentFolder, folderName);

			Marshal.ReleaseComObject(parentFolder);

			return folderExists;
		}

		private static IDictionary<string, IList<string>> GetFolderHashTable(
			string path, MAPIFolder folder)
		{
			IDictionary<string, IList<string>> hashTable = null;

			if (folder != null)
			{
				hashTable = new Dictionary<string, IList<string>>();
				Items items = folder.Items;
				int total = items.Count;

				Log.Info("Checking for duplicates at: " + path +
					" Total items: " + total);

				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group will change.
				for (int index = total; index > 0; index--)
				{
					object item = items[index];

					switch (item)
					{
						// Initially, just focus on MailItems
						case MailItem mailItem:
							string hash =
								MapiItem.GetItemHash(path, mailItem);

							if (!string.IsNullOrEmpty(hash))
							{
								bool keyExists = hashTable.ContainsKey(hash);

								if (keyExists == true)
								{
									IList<string> bucket = hashTable[hash];
									bucket.Add(mailItem.EntryID);
								}
								else
								{
									IList<string> bucket = new List<string>();
									bucket.Add(mailItem.EntryID);

									hashTable.Add(hash, bucket);
								}
							}

							Marshal.ReleaseComObject(mailItem);
							break;
						default:
							Log.Info("Ignoring item of non-MailItem type: ");
							break;
					}

					Marshal.ReleaseComObject(item);
				}
			}

			return hashTable;
		}

		private static string GetMailItemSynopses(MailItem mailItem)
		{
			string sentOn = mailItem.SentOn.ToString(
				"yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

			string synopses = string.Format(
				CultureInfo.InvariantCulture,
				"{0}: From: {1}: {2} Subject: {3}",
				sentOn,
				mailItem.SenderName,
				mailItem.SenderEmailAddress,
				mailItem.Subject);

			return synopses;
		}

		private static void ListItem(MailItem mailItem, string prefixMessage)
		{
			string sentOn = mailItem.SentOn.ToString(
				"yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

			string message = string.Format(
				CultureInfo.InvariantCulture,
				"{0} {1}: From: {2}: {3} Subject: {4}",
				prefixMessage,
				sentOn,
				mailItem.SenderName,
				mailItem.SenderEmailAddress,
				mailItem.Subject);

			Log.Info(message);
		}

		private static void MoveFolderItems(
			MAPIFolder source, MAPIFolder destination)
		{
			Items items = source.Items;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = items.Count; index > 0; index--)
			{
				object item = items[index];

				MapiItem.Moveitem(item, destination);
			}
		}

		private void CheckForDuplicateFolders(
			string path, int index, MAPIFolder folder)
		{
			string[] duplicatePatterns =
			{
				@"\s*\(\d*?\)", @"\s*-\s*Copy"
			};

			foreach (string duplicatePattern in duplicatePatterns)
			{
				MergeDuplicateFolder(path, index, folder, duplicatePattern);
			}
		}

		private int DeleteDuplicates(IList<string> duplicateSet, bool dryRun)
		{
			int totalDuplicates = duplicateSet.Count;

			string keeper = duplicateSet[0];
			duplicateSet.RemoveAt(0);

			NameSpace session = outlookAccount.Session;

			MailItem mailItem = session.GetItemFromID(keeper);
			string keeperSynopses = GetMailItemSynopses(mailItem);

			string message = string.Format(
				CultureInfo.InvariantCulture,
				"{0} Duplicates Found for: ",
				totalDuplicates.ToString(CultureInfo.InvariantCulture));

			ListItem(mailItem, message);

			foreach (string duplicateId in duplicateSet)
			{
				mailItem = session.GetItemFromID(duplicateId);
				string duplicateSynopses = GetMailItemSynopses(mailItem);

				if (!duplicateSynopses.Equals(
					keeperSynopses, StringComparison.Ordinal))
				{
					Log.Error("Warning! Duplicate Items Don't Seem to Match");
					Log.Error("Not Matching Item: " + duplicateSynopses);
				}

				if (dryRun == false)
				{
					mailItem.Delete();
				}

				Marshal.ReleaseComObject(mailItem);
			}

			return totalDuplicates;
		}

		private void MergeDuplicateFolder(
			string path, int index, MAPIFolder folder, string duplicatePattern)
		{
			if (Regex.IsMatch(
				folder.Name, duplicatePattern, RegexOptions.IgnoreCase))
			{
				string newFolderName = Regex.Replace(
					folder.Name,
					duplicatePattern,
					string.Empty,
					RegexOptions.ExplicitCapture);

				bool folderExists =
					DoesSiblingFolderExist(folder, newFolderName);

				if (folderExists == true)
				{
					MAPIFolder parentFolder = folder.Parent;

					// Move items
					MAPIFolder destination =
						parentFolder.Folders[newFolderName];

					MoveFolderContents(path, folder, destination);

					path += "/" + folder.Name;

					// Once all the items have been moved,
					// now remove the folder.
					RemoveFolder(path, index, folder, false);
				}
				else
				{
					try
					{
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

		private void MoveSubFolders(
			string path, MAPIFolder source, MAPIFolder destination)
		{
			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = source.Folders.Count; index > 0; index--)
			{
				MAPIFolder subFolder = source.Folders[index];

				MoveFolder(path, subFolder, destination, index);
			}
		}

		private void MoveFolder(
			string path, MAPIFolder source, MAPIFolder destination, int index)
		{
			string destinationName = destination.Name;

			string name = source.Name;
			MAPIFolder destinationSubFolder = GetSubFolder(destination, name);

			if (destinationSubFolder == null)
			{
				// Folder doesn't already exist, so just move it.
				string message = string.Format(
					CultureInfo.InvariantCulture,
					"at: {0} Moving {1} to {2}",
					path,
					name,
					destinationName);
				Log.Info(message);

				source.MoveTo(destination);
			}
			else
			{
				// Folder exists, so if just moving it, it will get
				// renamed something FolderName (2), so need to merge.
				string subPath = path + "/" + source.Name;

				string message = string.Format(
					CultureInfo.InvariantCulture,
					"at: {0} Merging {1} to {2}",
					subPath,
					name,
					destinationName);
				Log.Info(message);
				MoveFolderContents(
					subPath, source, destinationSubFolder);

				// Once all the items have been moved,
				// now remove the folder.
				RemoveFolder(subPath, index, source, false);
			}
		}

		private int[] RemoveDuplicatesFromThisFolder(
			MAPIFolder folder, bool dryRun)
		{
			int[] duplicateCounts = new int[2];

			string path = GetFolderPath(folder);

			IDictionary<string, IList<string>> hashTable =
				GetFolderHashTable(path, folder);

			var duplicates = hashTable.Where(p => p.Value.Count > 1);
			duplicateCounts[0] = duplicates.Count();

			if (duplicateCounts[0] > 0)
			{
				Log.Info("Duplicates found at: " + path);
			}

			foreach (KeyValuePair<string, IList<string>> duplicateSet in
				duplicates)
			{
				duplicateCounts[1] +=
					DeleteDuplicates(duplicateSet.Value, dryRun);
			}

			Marshal.ReleaseComObject(folder);

			return duplicateCounts;
		}
	}
}
