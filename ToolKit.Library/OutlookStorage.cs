/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookStorage.cs" company="James John McGuire">
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
	/// Provides support for interacting with an Outlook PST file.
	/// </summary>
	public class OutlookStorage
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private readonly Application outlookApplication;
		private readonly string[] ignoreFolders =
		{
				"Calendar", "Contacts", "Conversation Action Settings",
				"Deleted Items", "Deleted Messages", "Drafts", "Junk E-mail",
				"Journal", "Notes", "Outbox", "Quick Step Settings",
				"RSS Feeds", "Search Folders", "Sent Items", "Tasks"
		};

		private readonly NameSpace outlookNamespace;

		private uint totalFolders;
		private uint removedFolders;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookStorage"/> class.
		/// </summary>
		public OutlookStorage()
		{
			outlookApplication = new ();

			outlookNamespace = outlookApplication.GetNamespace("mapi");
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
				MAPIFolder parent = folder.Parent;

				while (parent is not null && parent is MAPIFolder)
				{
					path = parent.Name + "/" + path;
					folder = parent;

					if (folder.Parent is not null && folder.Parent is MAPIFolder)
					{
						parent = folder.Parent;
					}
					else
					{
						parent = null;
					}
				}

				string storeName = GetStoreName(folder.Store);
				path = storeName + "::" + path;
			}

			return path;
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
				int total = parentFolder.Folders.Count;

				for (int index = 1; index <= total; index++)
				{
					MAPIFolder subFolder = parentFolder.Folders[index];

					if (folderName.Equals(
						subFolder.Name, StringComparison.OrdinalIgnoreCase))
					{
						pstFolder = subFolder;
						break;
					}

					Marshal.ReleaseComObject(subFolder);
				}
			}

			return pstFolder;
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

				pstFolder = AddFolder(rootFolder, folderName);

				Marshal.ReleaseComObject(rootFolder);
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
					try
					{
						MailItem item =
							outlookNamespace.OpenSharedItem(filePath);

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
		/// Create mail item.
		/// </summary>
		/// <param name="recipient">The recipient of the mail.</param>
		/// <param name="subject">The subject of the mail.</param>
		/// <param name="body">The body of the mail.</param>
		/// <returns>The created mail item.</returns>
		public MailItem CreateMailItem(
			string recipient, string subject, string body)
		{
			MailItem mailItem =
				(MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);

			mailItem.Display(false);

			mailItem.To = recipient;
			mailItem.Subject = subject;
			mailItem.Body = body;

			return mailItem;
		}

		/// <summary>
		/// Create a new pst storage file.
		/// </summary>
		/// <param name="path">The path to the pst file.</param>
		/// <returns>A store object.</returns>
		public Store GetStore(string path)
		{
			Store newPst = null;

			// If the .pst file does not exist, Microsoft Outlook creates it.
			outlookNamespace.Session.AddStore(path);

			int total = outlookNamespace.Session.Stores.Count;

			for (int index = 1; index <= total; index++)
			{
				Store store = outlookNamespace.Session.Stores[index];

				if (store == null)
				{
					Log.Warn("Enumerating stores - store is null");
				}
				else
				{
					string filePath = store.FilePath;

					if (!string.IsNullOrWhiteSpace(filePath) &&
						store.FilePath == path)
					{
						newPst = store;
						break;
					}
				}
			}

			return newPst;
		}

		/// <summary>
		/// Empty deleted items folder.
		/// </summary>
		public void EmptyDeletedItemsFolder()
		{
			MAPIFolder deletedItemsFolder = outlookNamespace.GetDefaultFolder(
					OlDefaultFolders.olFolderDeletedItems);
			Items items = deletedItemsFolder.Items;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group will change.
			for (int index = items.Count; index > 0; index--)
			{
				object item = items[index];

				switch (item)
				{
					case AppointmentItem appointmentItem:
						appointmentItem.Delete();
						Marshal.ReleaseComObject(appointmentItem);
						break;
					case ContactItem contactItem:
						contactItem.Delete();
						Marshal.ReleaseComObject(contactItem);
						break;
					case DistListItem distListItem:
						distListItem.Delete();
						Marshal.ReleaseComObject(distListItem);
						break;
					case DocumentItem documentItem:
						documentItem.Delete();
						Marshal.ReleaseComObject(documentItem);
						break;
					case JournalItem journalItem:
						journalItem.Delete();
						Marshal.ReleaseComObject(journalItem);
						break;
					case MailItem mailItem:
						mailItem.Delete();
						Marshal.ReleaseComObject(mailItem);
						break;
					case MeetingItem meetingItem:
						meetingItem.Delete();
						Marshal.ReleaseComObject(meetingItem);
						break;
					case NoteItem noteItem:
						noteItem.Delete();
						Marshal.ReleaseComObject(noteItem);
						break;
					case PostItem postItem:
						postItem.Delete();
						Marshal.ReleaseComObject(postItem);
						break;
					case RemoteItem remoteItem:
						remoteItem.Delete();
						Marshal.ReleaseComObject(remoteItem);
						break;
					case ReportItem reportItem:
						reportItem.Delete();
						Marshal.ReleaseComObject(reportItem);
						break;
					case TaskItem taskItem:
						taskItem.Delete();
						Marshal.ReleaseComObject(taskItem);
						break;
					case TaskRequestAcceptItem taskRequestAcceptItem:
						taskRequestAcceptItem.Delete();
						Marshal.ReleaseComObject(taskRequestAcceptItem);
						break;
					case TaskRequestDeclineItem taskRequestDeclineItem:
						taskRequestDeclineItem.Delete();
						Marshal.ReleaseComObject(taskRequestDeclineItem);
						break;
					case TaskRequestItem taskRequestItem:
						taskRequestItem.Delete();
						Marshal.ReleaseComObject(taskRequestItem);
						break;
					case TaskRequestUpdateItem taskRequestUpdateItem:
						taskRequestUpdateItem.Delete();
						Marshal.ReleaseComObject(taskRequestUpdateItem);
						break;
					default:
						Log.Warn(
							"folder item of unknown type: " + item.ToString());
						break;
				}

				Marshal.ReleaseComObject(item);
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
				folder =
					outlookNamespace.GetFolderFromID(entryId, store.StoreID);
			}

			return folder;
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		public void MergeFolders()
		{
			int total = outlookNamespace.Session.Stores.Count;

			for (int index = 1; index <= total; index++)
			{
				Store store = outlookNamespace.Session.Stores[index];

				MergeFolders(store);
			}

			Log.Info("Remove empty folder complete - total folders checked: " +
				totalFolders);
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="pstFilePath">The PST file to check.</param>
		public void MergeFolders(string pstFilePath)
		{
			Store store = GetStore(pstFilePath);

			MergeFolders(store);

			Log.Info("Remove empty folder complete - total folders checked: " +
				totalFolders);
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
		/// Remove duplicates items from default account.
		/// </summary>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		public void RemoveDuplicates(bool dryRun)
		{
			int total = outlookNamespace.Session.Stores.Count;

			for (int index = 1; index <= total; index++)
			{
				Store store = outlookNamespace.Session.Stores[index];

				RemoveDuplicates(store, dryRun);
			}
		}

		/// <summary>
		/// Remove duplicates items from the given folder.
		/// </summary>
		/// <param name="folder">The MAPI folder to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <param name="recurse">Indicates whether to recurse into
		/// sub folders.</param>
		/// <returns>An array of duplicate sets and total duplicate items
		/// count.</returns>
		public int[] RemoveDuplicates(
			MAPIFolder folder, bool dryRun, bool recurse)
		{
			int[] duplicateCounts = new int[2];

			if (folder != null)
			{
				string folderName = folder.Name;

				if (!ignoreFolders.Contains(folderName))
				{
					if (recurse == true)
					{
						string path = GetFolderPath(folder);
						duplicateCounts = RemoveDuplicatesFromSubFolders(
							path, folder, dryRun);
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
		/// Remove duplicates items from the given store.
		/// </summary>
		/// <param name="storePath">The path of the PST file to
		/// process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		public void RemoveDuplicates(string storePath, bool dryRun)
		{
			Store store = GetStore(storePath);

			if (store != null)
			{
				RemoveDuplicates(store, dryRun);
				Marshal.ReleaseComObject(store);
			}
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		public void RemoveEmptyFolders()
		{
			int total = outlookNamespace.Session.Stores.Count;

			for (int index = 1; index <= total; index++)
			{
				Store store = outlookNamespace.Session.Stores[index];

				string path = store.FilePath;
				string extension = Path.GetExtension(path);

				if (extension.Equals(
					".ost", StringComparison.OrdinalIgnoreCase))
				{
					// for the time being, ignore ost files.
					continue;
				}

				string storePath = GetStoreName(store) + "::";

				MAPIFolder rootFolder = store.GetRootFolder();

				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group may change.
				for (int subIndex = rootFolder.Folders.Count; subIndex > 0;
					subIndex--)
				{
					path = storePath + rootFolder.Name;

					MAPIFolder subFolder = rootFolder.Folders[subIndex];
					bool subFolderEmtpy = RemoveEmptyFolders(path, subFolder);

					if (subFolderEmtpy == true)
					{
						string name = subFolder.Name;

						if (ignoreFolders.Contains(name))
						{
							Log.Warn("Not deleting reserved folder: " +
								name);
						}
						else
						{
							RemoveFolder(path, subIndex, subFolder, false);
						}
					}

					totalFolders++;
					Marshal.ReleaseComObject(subFolder);
				}

				totalFolders++;
				Marshal.ReleaseComObject(rootFolder);
			}

			Log.Info("Remove empty folder complete - total folder checked:" +
				totalFolders);
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="folder">The current folder.</param>
		/// <returns>Indicates whether the current folder is empty
		/// or not.</returns>
		public bool RemoveEmptyFolders(string path, MAPIFolder folder)
		{
			bool isEmpty = false;

			if (folder != null)
			{
				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group may change.
				for (int index = folder.Folders.Count; index > 0; index--)
				{
					MAPIFolder subFolder = folder.Folders[index];

					string subPath = path + "/" + subFolder.Name;

					bool subFolderEmtpy =
						RemoveEmptyFolders(subPath, subFolder);

					if (subFolderEmtpy == true)
					{
						RemoveFolder(subPath, index, subFolder, false);
					}

					totalFolders++;
					Marshal.ReleaseComObject(subFolder);
				}

				if (folder.Folders.Count == 0 && folder.Items.Count == 0)
				{
					isEmpty = true;
				}
			}

			return isEmpty;
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

				RemoveFolder(path, index, subFolder, force);
			}
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
					path += "/" + subFolder.Name;
					Log.Info("Removing empty folder: " + path);

					try
					{
						MAPIFolder parentFolder = subFolder.Parent;

						parentFolder.Folders.Remove(subFolderIndex);
					}
					catch (COMException exception)
					{
						Log.Error(exception.ToString());
					}

					removedFolders++;
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

				Marshal.ReleaseComObject(rootFolder);
				Marshal.ReleaseComObject(store);
			}
		}

		private static bool DoesSiblingFolderExist(
			MAPIFolder folder, string folderName)
		{
			bool folderExists = false;

			MAPIFolder parentFolder = folder.Parent;

			int total = parentFolder.Folders.Count;

			for (int index = 1; index <= total; index++)
			{
				MAPIFolder subFolder = parentFolder.Folders[index];

				string name = subFolder.Name;

				if (folderName.Equals(
					name, StringComparison.OrdinalIgnoreCase))
				{
					folderExists = true;
					break;
				}

				Marshal.ReleaseComObject(subFolder);
			}

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

				switch (item)
				{
					case AppointmentItem appointmentItem:
						appointmentItem.Move(destination);
						Marshal.ReleaseComObject(appointmentItem);
						break;
					case ContactItem contactItem:
						contactItem.Move(destination);
						Marshal.ReleaseComObject(contactItem);
						break;
					case DistListItem distListItem:
						distListItem.Move(destination);
						Marshal.ReleaseComObject(distListItem);
						break;
					case DocumentItem documentItem:
						documentItem.Move(destination);
						Marshal.ReleaseComObject(documentItem);
						break;
					case JournalItem journalItem:
						journalItem.Move(destination);
						Marshal.ReleaseComObject(journalItem);
						break;
					case MailItem mailItem:
						mailItem.Move(destination);
						Marshal.ReleaseComObject(mailItem);
						break;
					case MeetingItem meetingItem:
						meetingItem.Move(destination);
						Marshal.ReleaseComObject(meetingItem);
						break;
					case NoteItem noteItem:
						noteItem.Move(destination);
						Marshal.ReleaseComObject(noteItem);
						break;
					case PostItem postItem:
						postItem.Move(destination);
						Marshal.ReleaseComObject(postItem);
						break;
					case RemoteItem remoteItem:
						remoteItem.Move(destination);
						Marshal.ReleaseComObject(remoteItem);
						break;
					case ReportItem reportItem:
						reportItem.Move(destination);
						Marshal.ReleaseComObject(reportItem);
						break;
					case TaskItem taskItem:
						taskItem.Move(destination);
						Marshal.ReleaseComObject(taskItem);
						break;
					case TaskRequestAcceptItem taskRequestAcceptItem:
						taskRequestAcceptItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestAcceptItem);
						break;
					case TaskRequestDeclineItem taskRequestDeclineItem:
						taskRequestDeclineItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestDeclineItem);
						break;
					case TaskRequestItem taskRequestItem:
						taskRequestItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestItem);
						break;
					case TaskRequestUpdateItem taskRequestUpdateItem:
						taskRequestUpdateItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestUpdateItem);
						break;
					default:
						Log.Warn(
							"folder item of unknown type: " + item.ToString());
						break;
				}

				Marshal.ReleaseComObject(item);
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

			MailItem mailItem = outlookNamespace.GetItemFromID(keeper);
			string keeperSynopses = GetMailItemSynopses(mailItem);

			string message = string.Format(
				CultureInfo.InvariantCulture,
				"{0} Duplicates Found for: ",
				totalDuplicates.ToString(CultureInfo.InvariantCulture));

			ListItem(mailItem, message);

			foreach (string duplicateId in duplicateSet)
			{
				mailItem = outlookNamespace.GetItemFromID(duplicateId);
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

		private void MoveFolderContents(
			string path, MAPIFolder source, MAPIFolder destination)
		{
			MoveFolderItems(source, destination);
			MoveFolderFolders(path, source, destination);
		}

		private void MoveFolderFolders(
			string path, MAPIFolder source, MAPIFolder destination)
		{
			string destinationName = destination.Name;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = source.Folders.Count; index > 0; index--)
			{
				MAPIFolder subFolder = source.Folders[index];

				string name = subFolder.Name;
				MAPIFolder destinationSubFolder =
					GetSubFolder(destination, name);

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
					subFolder.MoveTo(destination);
				}
				else
				{
					// Folder exists, so if just moving it, it will get
					// renamed something FolderName (2), so need to merge.
					string subPath = path + "/" + subFolder.Name;

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"at: {0} Merging {1} to {2}",
						subPath,
						name,
						destinationName);
					Log.Info(message);
					MoveFolderContents(
						subPath, subFolder, destinationSubFolder);

					// Once all the items have been moved,
					// now remove the folder.
					RemoveFolder(path, index, subFolder, false);
				}
			}
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

		private void MergeFolders(Store store)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store) + "::";
				MAPIFolder rootFolder = store.GetRootFolder();

				MergeFolders(storePath, rootFolder);

				totalFolders++;

				Marshal.ReleaseComObject(rootFolder);
				Marshal.ReleaseComObject(store);
			}
		}

		/// <summary>
		/// Remove duplicates items from the given store.
		/// </summary>
		/// <param name="store">The PST store to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		private void RemoveDuplicates(Store store, bool dryRun)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store);
				Log.Info("Checking for duplicates in: " + storePath);

				MAPIFolder rootFolder = store.GetRootFolder();

				int[] duplicateCounts = RemoveDuplicatesFromSubFolders(
					storePath, rootFolder, dryRun);

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
		/// Remove duplicates items from the given folder.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="folder">The MAPI folder to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <param name="recurse">Indicates whether to recurse into
		/// sub folders.</param>
		/// <returns>An array of duplicate sets and total duplicate items
		/// count.</returns>
		private int[] RemoveDuplicates(
			string path, MAPIFolder folder, bool dryRun, bool recurse)
		{
			int[] duplicateCounts = new int[2];

			if (!ignoreFolders.Contains(folder.Name))
			{
				if (recurse == true)
				{
					duplicateCounts =
						RemoveDuplicatesFromSubFolders(path, folder, dryRun);
				}

				int[] duplicateCountsThisFolder =
					RemoveDuplicatesFromThisFolder(folder, dryRun);

				duplicateCounts[0] += duplicateCountsThisFolder[0];
				duplicateCounts[1] += duplicateCountsThisFolder[1];
			}

			return duplicateCounts;
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
		private int[] RemoveDuplicatesFromSubFolders(
			string path, MAPIFolder folder, bool dryRun)
		{
			int[] duplicateCounts = new int[2];

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = folder.Folders.Count; index > 0; index--)
			{
				MAPIFolder subFolder = folder.Folders[index];

				int[] subFolderduplicateCounts =
					RemoveDuplicates(path, subFolder, dryRun, true);

				duplicateCounts[0] += subFolderduplicateCounts[0];
				duplicateCounts[1] += subFolderduplicateCounts[1];

				totalFolders++;
				Marshal.ReleaseComObject(subFolder);
			}

			return duplicateCounts;
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
