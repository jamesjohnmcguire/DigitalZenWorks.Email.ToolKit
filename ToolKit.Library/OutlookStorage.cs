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

		private readonly OutlookAccount outlookAccount;
		private readonly string[] ignoreFolders =
		{
				"Calendar", "Contacts", "Conversation Action Settings",
				"Deleted Items", "Deleted Messages", "Drafts", "Junk E-mail",
				"Journal", "Notes", "Outbox", "Quick Step Settings",
				"RSS Feeds", "Search Folders", "Sent Items", "Tasks"
		};

		private uint totalFolders;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookStorage"/> class.
		/// </summary>
		/// <param name="outlookAccount">The outlook account object.</param>
		public OutlookStorage(OutlookAccount outlookAccount)
		{
			this.outlookAccount = outlookAccount;
		}

		/// <summary>
		/// Gets the ignore folders list.
		/// </summary>
		/// <value>The ignore folders list.</value>
		public string[] IgnoreFolders { get { return ignoreFolders; } }

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
		public static void RemoveFolder(
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

				OutlookFolder outlookFolder = new ();
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

			Log.Info("Remove empty folder complete - total folders checked: " +
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

				OutlookFolder outlookFolder = new ();
				outlookFolder.MergeFolders(storePath, rootFolder);

				totalFolders++;

				Marshal.ReleaseComObject(rootFolder);
				Marshal.ReleaseComObject(store);
			}

			return totalFolders;
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
			Store store = outlookAccount.GetStore(storePath);

			if (store != null)
			{
				RemoveDuplicates(store, dryRun);
				Marshal.ReleaseComObject(store);
			}
		}

		/// <summary>
		/// Remove duplicates items from the given store.
		/// </summary>
		/// <param name="store">The PST store to process.</param>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		public void RemoveDuplicates(Store store, bool dryRun)
		{
			if (store != null)
			{
				string storePath = GetStoreName(store);
				Log.Info("Checking for duplicates in: " + storePath);

				MAPIFolder rootFolder = store.GetRootFolder();

				OutlookFolder outlookFolder = new ();
				int[] duplicateCounts =
					outlookFolder.RemoveDuplicatesFromSubFolders(
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
		/// Remove all empty folders.
		/// </summary>
		/// <param name="store">The PST store to process.</param>
		public uint RemoveEmptyFolders(Store store)
		{
			if (store != null)
			{
				string path = store.FilePath;
				string extension = Path.GetExtension(path);

				if (!extension.Equals(
					".ost", StringComparison.OrdinalIgnoreCase))
				{
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
								OutlookFolder outlookFolder = new ();
								outlookFolder.RemoveFolder(
									path, subIndex, subFolder, false);
							}
						}

						totalFolders++;
						Marshal.ReleaseComObject(subFolder);
					}

					totalFolders++;
					Marshal.ReleaseComObject(rootFolder);
				}
			}

			Log.Info("Remove empty folder complete - total folder checked:" +
				totalFolders);

			return totalFolders;
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
						OutlookFolder outlookFolder = new ();
						outlookFolder.RemoveFolder(
							subPath, index, subFolder, false);
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
	}
}
