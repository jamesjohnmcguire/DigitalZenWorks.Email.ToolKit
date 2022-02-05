/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookStorage.cs" company="James John McGuire">
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
		private readonly NameSpace outlookNamespace;

		private uint totalFolders;
		private uint removedFolders;

		/// <summary>
		/// Initializes a new instance of the <see cref="OutlookStorage"/> class.
		/// </summary>
		public OutlookStorage()
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
					name = Path.GetFileNameWithoutExtension(store.FilePath);
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
				foreach (MAPIFolder subFolder in parentFolder.Folders)
				{
					if (folderName.Equals(
						subFolder.Name, StringComparison.OrdinalIgnoreCase))
					{
						pstFolder = subFolder;
						break;
					}
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

			mailItem.To = recipient;
			mailItem.Subject = subject;
			mailItem.Body = body;
			mailItem.Display(false);

			return mailItem;
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
		/// Merge duplicate folders.
		/// </summary>
		public void MergeFolders()
		{
			foreach (Store store in outlookNamespace.Session.Stores)
			{
				string storePath = GetStoreName(store) + "::";

				MAPIFolder rootFolder = store.GetRootFolder();

				// Office uses 1 based indexes from VBA.
				// Iterate in reverse order as the group may change.
				for (int index = rootFolder.Folders.Count; index > 0; index--)
				{
					string path = storePath + rootFolder.Name;

					MAPIFolder subFolder = rootFolder.Folders[index];
					MergeFolders(path, subFolder);

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
		/// Remove all empty folders.
		/// </summary>
		public void RemoveEmptyFolders()
		{
			string[] ignoreFolders =
			{
				"Calendar", "Contacts", "Conversation Action Settings",
				"Deleted Items", "Drafts", "Junk E-mail", "Journal", "Notes",
				"Outbox", "Quick Step Settings", "RSS Feeds", "Search Folders",
				"Sent Items", "Tasks"
			};

			foreach (Store store in outlookNamespace.Session.Stores)
			{
				string extension = Path.GetExtension(store.FilePath);

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
				for (int index = rootFolder.Folders.Count; index > 0; index--)
				{
					string path = storePath + rootFolder.Name;

					MAPIFolder subFolder = rootFolder.Folders[index];
					bool subFolderEmtpy = RemoveEmptyFolders(path, subFolder);

					if (subFolderEmtpy == true)
					{
						if (ignoreFolders.Contains(subFolder.Name))
						{
							Log.Warn("Not deleting reserved folder: " +
								subFolder.Name);
						}
						else
						{
							RemoveFolder(path, index, subFolder, false);
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
				MAPIFolder parentFolder = subFolder.Parent;

				int count = parentFolder.Folders.Count;
				int index;
				for (index = 1; index <= count; index++)
				{
					MAPIFolder folder = parentFolder.Folders[index];

					if (folder.Name.Equals(
						subFolder.Name, StringComparison.OrdinalIgnoreCase))
					{
						break;
					}
				}

				RemoveFolder(path, index, subFolder, false);
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
			}
		}

		private static bool DoesSiblingFolderExist(
					MAPIFolder folder, string folderName)
		{
			bool folderExists = false;

			MAPIFolder parentFolder = folder.Parent;

			foreach (MAPIFolder subFolder in parentFolder.Folders)
			{
				if (folderName.Equals(
					subFolder.Name, StringComparison.OrdinalIgnoreCase))
				{
					folderExists = true;
					break;
				}
			}

			Marshal.ReleaseComObject(parentFolder);

			return folderExists;
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
						break;
					case ContactItem contactItem:
						contactItem.Move(destination);
						break;
					case DistListItem distListItem:
						distListItem.Move(destination);
						break;
					case DocumentItem documentItem:
						documentItem.Move(destination);
						break;
					case JournalItem journalItem:
						journalItem.Move(destination);
						break;
					case MailItem mailItem:
						mailItem.Move(destination);
						Marshal.ReleaseComObject(mailItem);
						break;
					case MeetingItem meetingItem:
						meetingItem.Move(destination);
						break;
					case NoteItem noteItem:
						noteItem.Move(destination);
						break;
					case PostItem postItem:
						postItem.Move(destination);
						break;
					case RemoteItem remoteItem:
						remoteItem.Move(destination);
						break;
					case ReportItem reportItem:
						reportItem.Move(destination);
						break;
					case TaskItem taskItem:
						taskItem.Move(destination);
						break;
					case TaskRequestAcceptItem taskRequestAcceptItem:
						taskRequestAcceptItem.Move(destination);
						break;
					case TaskRequestDeclineItem taskRequestDeclineItem:
						taskRequestDeclineItem.Move(destination);
						break;
					case TaskRequestItem taskRequestItem:
						taskRequestItem.Move(destination);
						break;
					case TaskRequestUpdateItem taskRequestUpdateItem:
						taskRequestUpdateItem.Move(destination);
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

		private void MoveFolderContents(
			string path, MAPIFolder source, MAPIFolder destination)
		{
			MoveFolderItems(source, destination);
			MoveFolderFolders(path, source, destination);
		}

		private void MoveFolderFolders(
			string path, MAPIFolder source, MAPIFolder destination)
		{
			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = source.Folders.Count; index > 0; index--)
			{
				MAPIFolder subFolder = source.Folders[index];

				MAPIFolder destinationSubFolder =
					GetSubFolder(destination, subFolder.Name);

				if (destinationSubFolder == null)
				{
					// Folder doesn't already exist, so just move it.
					string message = string.Format(
						CultureInfo.InvariantCulture,
						"at: {0} Moving {1} to {2}",
						path,
						subFolder.Name,
						destination.Name);
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
						subFolder.Name,
						destination.Name);
					Log.Info(message);
					MoveFolderContents(subPath, subFolder, destinationSubFolder);

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
	}
}
