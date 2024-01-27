/////////////////////////////////////////////////////////////////////////////
// <copyright file="Migrate.cs" company="James John McGuire">
// Copyright © 2021 - 2024 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Email.DbxOutlookExpress;
using Microsoft.Office.Interop.Outlook;
using MsgKit;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

[assembly: CLSCompliant(false)]

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Migrate Dbx to Pst class.
	/// </summary>
	public static class Migrate
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// Dbx directory to pst.
		/// </summary>
		/// <param name="dbxFoldersPath">The path to dbx folders to
		/// migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		public static void DbxDirectoryToPst(
			string dbxFoldersPath, string pstPath)
		{
			DbxDirectoryToPst(dbxFoldersPath, pstPath, null);
		}

		/// <summary>
		/// Dbx directory to pst.
		/// </summary>
		/// <param name="dbxFoldersPath">The path to dbx folders to
		/// migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <param name="encoding">The optional encoding to use.</param>
		public static void DbxDirectoryToPst(
			string dbxFoldersPath, string pstPath, Encoding encoding)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;

			Store pstStore = outlookAccount.GetStore(pstPath);

			if (pstStore == null)
			{
				Log.Error("PST store not created");
			}
			else
			{
				DbxSet dbxSet = new (dbxFoldersPath, encoding);

				// Order the list, so that parents always come before their
				// children.
				dbxSet.SetTreeOrdered();

				DbxFolder dbxFolder;
				MAPIFolder rootFolder = pstStore.GetRootFolder();

				string baseName = Path.GetFileNameWithoutExtension(pstPath);
				rootFolder.Name = baseName;

				IDictionary<uint, string> mappings =
					new Dictionary<uint, string>();

				do
				{
					dbxFolder = dbxSet.GetNextFolder();

					CopyFolderToPst(
						mappings,
						pstStore,
						rootFolder,
						dbxFolder);
				}
				while (dbxFolder != null);

				Marshal.ReleaseComObject(rootFolder);
			}
		}

		/// <summary>
		/// Dbx files to pst.
		/// </summary>
		/// <param name="filePath">The file path to migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		public static void DbxFileToPst(string filePath, string pstPath)
		{
			DbxFileToPst(filePath, pstPath, null);
		}

		/// <summary>
		/// Dbx files to pst.
		/// </summary>
		/// <param name="filePath">The file path to migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <param name="encoding">The optional encoding to use.</param>
		public static void DbxFileToPst(
			string filePath, string pstPath, Encoding encoding)
		{
			FileInfo fileInfo = new (filePath);

			if (fileInfo.Name.Equals(
				"Folders.dbx", StringComparison.OrdinalIgnoreCase))
			{
				DbxDirectoryToPst(filePath, pstPath, encoding);
			}
			else
			{
				OutlookAccount outlookAccount = OutlookAccount.Instance;
				Store pstStore = outlookAccount.GetStore(pstPath);

				MAPIFolder rootFolder = pstStore.GetRootFolder();

				string baseName = Path.GetFileNameWithoutExtension(pstPath);
				rootFolder.Name = baseName;

				DbxFolder dbxFolder = new (filePath, baseName, encoding);

				MAPIFolder pstFolder = OutlookFolder.AddFolder(
					rootFolder, rootFolder.Name);

				CopyMessages(pstFolder, dbxFolder);

				Marshal.ReleaseComObject(pstFolder);
				Marshal.ReleaseComObject(rootFolder);
			}
		}

		/// <summary>
		/// Dbx to pst.
		/// </summary>
		/// <param name="path">the path of the dbx element.</param>
		/// <returns>A value indicating success or not.</returns>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		public static bool DbxToPst(string path, string pstPath)
		{
			return DbxToPst(path, pstPath, null);
		}

		/// <summary>
		/// Dbx to pst.
		/// </summary>
		/// <param name="path">the path of the dbx element.</param>
		/// <returns>A value indicating success or not.</returns>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <param name="encoding">The optional encoding to use.</param>
		public static bool DbxToPst(
			string path, string pstPath, Encoding encoding)
		{
			bool result = false;

			Log.Info("Checking file: " + path);

			if (Directory.Exists(path))
			{
				DbxDirectoryToPst(path, pstPath, encoding);
				result = true;
			}
			else if (File.Exists(path))
			{
				DbxFileToPst(path, pstPath, encoding);
				result = true;
			}
			else
			{
				Log.Error("Invalid path");
			}

			return result;
		}

		/// <summary>
		/// Eml file to pst.
		/// </summary>
		/// <remarks>The caller is responsible for deleting
		/// the object.</remarks>
		/// <param name="filePath">The file path to migrate.</param>
		/// <param name="folder">The Outlook folder to copy to.</param>
		/// <returns>A valid MailItem or null.</returns>
		public static MailItem EmlFileToPst(string filePath, MAPIFolder folder)
		{
			MailItem mailItem = null;

			if (folder != null)
			{
				try
				{
					mailItem = CopyEmlToPst(folder, filePath);
				}
				catch (IOException exception)
				{
					Log.Error(exception.ToString());
				}
			}

			return mailItem;
		}

		/// <summary>
		/// Eml file to pst.
		/// </summary>
		/// <remarks>The caller is responsible for deleting
		/// the object.</remarks>
		/// <param name="filePath">The file path to migrate.</param>
		/// <param name="pstStore">The pst store to copy to.</param>
		/// <param name="folderName">The name of the folder to copy to.</param>
		/// <returns>A valid MailItem or null.</returns>
		public static MailItem EmlFileToPst(
			string filePath, Store pstStore, string folderName)
		{
			MailItem mailItem = null;

			MAPIFolder pstFolder =
				OutlookStore.GetTopLevelFolder(pstStore, folderName);

			if (pstFolder != null)
			{
				mailItem = EmlFileToPst(filePath, pstFolder);

				Marshal.ReleaseComObject(pstFolder);
			}

			return mailItem;
		}

		/// <summary>
		/// Eml file to pst.
		/// </summary>
		/// <remarks>The caller is responsible for deleting
		/// the object.</remarks>
		/// <param name="filePath">The file path to migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <returns>A valid MailItem or null.</returns>
		public static MailItem EmlFileToPst(string filePath, string pstPath)
		{
			MailItem mailItem = null;

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			Store pstStore = outlookAccount.GetStore(pstPath);

			string baseName = Path.GetFileNameWithoutExtension(pstPath);

			MAPIFolder pstFolder =
				OutlookStore.GetTopLevelFolder(pstStore, baseName);

			if (pstFolder != null)
			{
				try
				{
					mailItem = CopyEmlToPst(pstFolder, filePath);
				}
				catch (IOException exception)
				{
					Log.Error(exception.ToString());
				}

				Marshal.ReleaseComObject(pstFolder);
			}

			return mailItem;
		}

		/// <summary>
		/// Eml file to pst.
		/// </summary>
		/// <remarks>The caller is responsible for deleting
		/// the object.</remarks>
		/// <param name="filePath">The file path to migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <param name="folderName">The name of the folder to copy to.</param>
		/// <returns>A valid MailItem or null.</returns>
		public static MailItem EmlFileToPst(
			string filePath, string pstPath, string folderName)
		{
			MailItem mailItem = null;

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			Store pstStore = outlookAccount.GetStore(pstPath);

			MAPIFolder pstFolder =
				OutlookStore.GetTopLevelFolder(pstStore, folderName);

			if (pstFolder != null)
			{
				mailItem = EmlFileToPst(filePath, pstFolder);

				Marshal.ReleaseComObject(pstFolder);
			}

			return mailItem;
		}

		/// <summary>
		/// Eml to pst.
		/// </summary>
		/// <param name="path">the path of the eml directory or file.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <returns>A value indicating success or not.</returns>
		public static bool EmlToPst(string path, string pstPath)
		{
			bool result = EmlToPst(path, pstPath, true);

			return result;
		}

		/// <summary>
		/// Eml to pst.
		/// </summary>
		/// <param name="path">the path of the eml directory or file.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <param name="adjust">Indicates whether to exclude interim
		/// folders.</param>
		/// <returns>A value indicating success or not.</returns>
		public static bool EmlToPst(string path, string pstPath, bool adjust)
		{
			bool result = false;

			Log.Info("Checking file: " + path);

			if (Directory.Exists(path))
			{
				OutlookAccount outlookAccount = OutlookAccount.Instance;
				Store store = outlookAccount.GetStore(pstPath);

				if (store == null)
				{
					Log.Error("PST store not created");
				}
				else
				{
					MAPIFolder rootFolder = store.GetRootFolder();

					string baseName =
						Path.GetFileNameWithoutExtension(pstPath);
					rootFolder.Name = baseName;

					EmlDirectoryToPst(rootFolder, path, adjust);

					Marshal.ReleaseComObject(rootFolder);
					result = true;
				}
			}
			else if (File.Exists(path))
			{
				EmlFileToPst(path, pstPath);
				result = true;
			}
			else
			{
				Log.Error("Invalid path");
			}

			return result;
		}

		private static void AddMappingSafe(
			IDictionary<uint, string> mappings,
			MAPIFolder pstFolder,
			DbxFolder dbxFolder)
		{
			bool keyExists =
				mappings.ContainsKey(dbxFolder.FolderId);

			if (keyExists == true)
			{
				LogFormatMessage.Info(
					"Duplicate key mapping! Folder Id: {1} Name: {0} ",
					dbxFolder.FolderName,
					dbxFolder.FolderId.ToString(CultureInfo.InvariantCulture));
			}
			else
			{
				mappings.Add(dbxFolder.FolderId, pstFolder.EntryID);
			}
		}

		private static MAPIFolder CopyChildFolderToPst(
			IDictionary<uint, string> mappings,
			OutlookStore pstOutlook,
			Store pstStore,
			DbxFolder dbxFolder)
		{
			MAPIFolder parentFolder;
			MAPIFolder pstFolder;

			// need to figure out parent in pst
			bool keyExists =
				mappings.ContainsKey(dbxFolder.FolderParentId);

			if (keyExists == false)
			{
				Log.Warn("Parent key not found in mappings: " +
					dbxFolder.FolderParentId);

				parentFolder = pstStore.GetRootFolder();
			}
			else
			{
				string entryId = mappings[dbxFolder.FolderParentId];
				parentFolder = pstOutlook.GetFolderFromID(entryId, pstStore);
			}

			pstFolder = OutlookFolder.AddFolder(
				parentFolder, dbxFolder.FolderName);

			Marshal.ReleaseComObject(parentFolder);

			return pstFolder;
		}

		private static void CopyEmlFilesToPst(
			MAPIFolder pstFolder, IReadOnlyCollection<string> emlFiles)
		{
			if (emlFiles.Count > 0)
			{
				foreach (string file in emlFiles)
				{
					try
					{
						CopyEmlToPst(pstFolder, file);
					}
					catch (IOException exception)
					{
						Log.Error(exception.ToString());
					}
				}
			}
		}

		private static MailItem CopyEmlToPst(MAPIFolder mapiFolder, string emlFile)
		{
			MailItem mailItem = null;

			if (!string.IsNullOrWhiteSpace(emlFile) && File.Exists(emlFile))
			{
				string msgFile = GetTemporaryFileName(".msg");

				try
				{
					Converter.ConvertEmlToMsg(emlFile, msgFile);
				}
				catch (IOException exception)
				{
					Log.Warn(exception.ToString());

					// Hmmmn, try one more time.
					msgFile = GetTemporaryFileName(".msg");
					Converter.ConvertEmlToMsg(emlFile, msgFile);
				}
				catch (System.Exception exception) when
					(exception is InvalidCastException ||
					exception is NullReferenceException)
				{
					Log.Error(exception.ToString());
				}

				OutlookAccount outlookAccount = OutlookAccount.Instance;
				OutlookFolder outlookFolder = new (outlookAccount);

				mailItem = outlookFolder.AddMsgFile(mapiFolder, msgFile);

				File.Delete(msgFile);
			}

			return mailItem;
		}

		/// <summary>
		/// Copy folder to pst store.
		/// </summary>
		/// <param name="mappings">The mappings file to add to.</param>
		/// <param name="pstStore">The store to use.</param>
		/// <param name="rootFolder">The root folder of the store.</param>
		/// <param name="dbxFolder">The dbx folder to add.</param>
		private static void CopyFolderToPst(
			IDictionary<uint, string> mappings,
			Store pstStore,
			MAPIFolder rootFolder,
			DbxFolder dbxFolder)
		{
			if (mappings != null && pstStore != null && dbxFolder != null)
			{
				MAPIFolder pstFolder;

				// The search folder doesn't seem to contain any actual
				// message content, so it would be justa a waste of time.
				if (!string.IsNullOrWhiteSpace(dbxFolder.FolderName) &&
					!dbxFolder.FolderName.Equals(
					"Search Folder", StringComparison.OrdinalIgnoreCase))
				{
					// add folder to pst
					if (dbxFolder.FolderParentId == 0)
					{
						if (dbxFolder.IsOrphan == true)
						{
							MAPIFolder orphanFolders =
								OutlookFolder.AddFolder(
									rootFolder, "Orphan Folders");
							pstFolder = OutlookFolder.AddFolder(
								orphanFolders, dbxFolder.FolderName);
						}
						else
						{
							// top level folder
							pstFolder = OutlookFolder.AddFolder(
								rootFolder, dbxFolder.FolderName);
						}
					}
					else
					{
						OutlookAccount outlookAccount =
							OutlookAccount.Instance;
						OutlookStore store = new (outlookAccount);

						pstFolder = CopyChildFolderToPst(
							mappings,
							store,
							pstStore,
							dbxFolder);
					}

					if (pstFolder != null)
					{
						AddMappingSafe(mappings, pstFolder, dbxFolder);

						CopyMessages(pstFolder, dbxFolder);
					}
					else
					{
						Log.Warn("pstFolder is null: " + dbxFolder.FolderName);
					}
				}
			}
		}

		private static void CopyMessages(
			MAPIFolder pstFolder,
			DbxFolder dbxFolder)
		{
			// for each message
			DbxMessage dbxMessage;

			do
			{
				dbxMessage = dbxFolder.GetNextMessage();

				CopyMessageToPst(pstFolder, dbxMessage);
			}
			while (dbxMessage != null);

			Marshal.ReleaseComObject(pstFolder);
		}

		private static void CopyMessageToPst(
			MAPIFolder mapiFolder,
			DbxMessage dbxMessage)
		{
			if (dbxMessage != null && dbxMessage.Message.Length > 0)
			{
				try
				{
					string filePath = GetTemporaryFileName(".eml");
					dbxMessage.GetAsFile(filePath);

					string msgFile = GetTemporaryFileName(".msg");

					Converter.ConvertEmlToMsg(filePath, msgFile);

					File.Delete(filePath);

					OutlookAccount outlookAccount = OutlookAccount.Instance;
					OutlookFolder outlookFolder =
						new (outlookAccount, mapiFolder);
					outlookFolder.AddMsgFile(mapiFolder, msgFile);

					File.Delete(msgFile);
				}
				catch (System.Exception exception) when
					(exception is ArgumentException ||
					exception is ArgumentNullException ||
					exception is DirectoryNotFoundException ||
					exception is FormatException ||
					exception is InvalidCastException ||
					exception is IOException ||
					exception is NotSupportedException ||
					exception is NullReferenceException ||
					exception is PathTooLongException ||
					exception is UnauthorizedAccessException)
				{
					Log.Error(exception.ToString());
				}
				}
			}

		private static void EmlDirectoryToPst(
			MAPIFolder pstParent, string emlFolderFilePath, bool adjust)
		{
			string[] directories = Directory.GetDirectories(emlFolderFilePath);

			IReadOnlyCollection<string> emlFiles =
				EmlMessages.GetFilesCollection(emlFolderFilePath);

			if (directories.Length > 0 || emlFiles.Count > 0)
			{
				try
				{
					DirectoryInfo directoryInfo = new (emlFolderFilePath);
					string directoryName = directoryInfo.Name;

					MAPIFolder thisFolder = pstParent;

					bool isInterimFolder =
						CheckIfInterimFolder(pstParent, directoryName);

					if (adjust == false || isInterimFolder == false)
					{
						thisFolder = OutlookFolder.AddFolder(
							pstParent, directoryName, true);
					}

					foreach (string directory in directories)
					{
						EmlDirectoryToPst(thisFolder, directory, adjust);
					}

					CopyEmlFilesToPst(thisFolder, emlFiles);

					if (adjust == false || isInterimFolder == false)
					{
						Marshal.ReleaseComObject(thisFolder);
					}
				}
				catch (InvalidComObjectException exception)
				{
					Log.Error(exception.ToString());
				}
			}
		}

		private static bool CheckIfInterimFolder(
			MAPIFolder pstParent, string currentDirectory)
		{
			bool interimFolder = false;

			bool isRoot = OutlookFolder.IsRootFolder(pstParent);

			if (isRoot == true &&
				(currentDirectory.Equals(
					"Local Folders", StringComparison.OrdinalIgnoreCase) ||
				currentDirectory.Equals(
					"Storage Folders", StringComparison.OrdinalIgnoreCase) ||
				currentDirectory.StartsWith(
					"Imported Fo", StringComparison.OrdinalIgnoreCase)))
			{
				interimFolder = true;
			}

			return interimFolder;
		}

		private static string GetTemporaryFileName(string extension)
		{
			string filePath = Path.GetTempFileName();

			// A 0 byte sized file is created.  Need to remove it.
			File.Delete(filePath);

			filePath = Path.ChangeExtension(filePath, extension);

			return filePath;
		}

		private static Stream GetTemporaryMsgFileStream(string msgFile)
		{
			Stream msgStream = new FileStream(msgFile, FileMode.Create);

			return msgStream;
		}
	}
}
