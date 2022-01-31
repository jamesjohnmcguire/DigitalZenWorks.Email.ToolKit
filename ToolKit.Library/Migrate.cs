﻿/////////////////////////////////////////////////////////////////////////////
// <copyright file="Migrate.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
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
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using ToolKit.Library;

[assembly: CLSCompliant(false)]

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// The transfer folder call back type.
	/// </summary>
	/// <param name="id">The folder id.</param>
	public delegate void TransferFolderCallBackType(int id);

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
			// Personal preference... For me, most of these types will
			// likely be Japansese.
			Encoding.RegisterProvider(
				CodePagesEncodingProvider.Instance);
			Encoding encoding = Encoding.GetEncoding("shift_jis");

			DbxSet dbxSet = new (dbxFoldersPath, encoding);

			// Order the list, so that parents always come before their
			// children.
			dbxSet.SetTreeOrdered();

			PstOutlook pstOutlook = new ();
			Store pstStore = pstOutlook.CreateStore(pstPath);

			if (pstStore == null)
			{
				Log.Error("PST store not created");
			}
			else
			{
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
						mappings, pstOutlook, pstStore, rootFolder, dbxFolder);
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
			// Personal preference... For me, most of these types will
			// likely be Japansese.
			Encoding.RegisterProvider(
				CodePagesEncodingProvider.Instance);
			Encoding encoding = Encoding.GetEncoding("shift_jis");

			FileInfo fileInfo = new (filePath);

			if (fileInfo.Name.Equals(
				"Folders.dbx", StringComparison.OrdinalIgnoreCase))
			{
				DbxDirectoryToPst(filePath, pstPath);
			}
			else
			{
				PstOutlook pstOutlook = new ();
				Store pstStore = pstOutlook.CreateStore(pstPath);

				MAPIFolder rootFolder = pstStore.GetRootFolder();

				string baseName = Path.GetFileNameWithoutExtension(pstPath);
				rootFolder.Name = baseName;

				DbxFolder dbxFolder = new (filePath, baseName, encoding);

				MAPIFolder pstFolder = PstOutlook.AddFolderSafe(
					rootFolder, rootFolder.Name);

				CopyMessages(pstOutlook, pstFolder, dbxFolder);

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
			bool result = false;

			Log.Info("Checking file: " + path);

			if (Directory.Exists(path))
			{
				DbxDirectoryToPst(path, pstPath);
				result = true;
			}
			else if (File.Exists(path))
			{
				DbxFileToPst(path, pstPath);
				result = true;
			}
			else
			{
				Log.Error("Invalid path");
			}

			return result;
		}

		/// <summary>
		/// Copy folder to pst store.
		/// </summary>
		/// <param name="mappings">The mappings file to add to.</param>
		/// <param name="pstOutlook">The pst object to use.</param>
		/// <param name="pstStore">The store to use.</param>
		/// <param name="rootFolder">The root folder of the store.</param>
		/// <param name="dbxFolder">The dbx folder to add.</param>
		public static void CopyFolderToPst(
			IDictionary<uint, string> mappings,
			PstOutlook pstOutlook,
			Store pstStore,
			MAPIFolder rootFolder,
			DbxFolder dbxFolder)
		{
			if (mappings != null && pstOutlook != null &&
				pstStore != null && dbxFolder != null)
			{
				MAPIFolder pstFolder;

				// The search folder doesn't seem to contain any actual
				// message content, so it would be justa a waste of time.
				if (!dbxFolder.FolderName.Equals(
					"Search Folder", StringComparison.OrdinalIgnoreCase))
				{
					// add folder to pst
					if (dbxFolder.FolderParentId == 0)
					{
						// top level folder
						pstFolder = PstOutlook.AddFolderSafe(
							rootFolder, dbxFolder.FolderName);
					}
					else
					{
						pstFolder = CopyChildFolderToPst(
							mappings,
							pstOutlook,
							pstStore,
							dbxFolder);
					}

					if (pstFolder != null)
					{
						AddMappingSafe(mappings, pstFolder, dbxFolder);

						CopyMessages(pstOutlook, pstFolder, dbxFolder);
					}
					else
					{
						Log.Warn("pstFolder is null: " + dbxFolder.FolderName);
					}
				}
			}
		}

		/// <summary>
		/// Dbx to pst.
		/// </summary>
		/// <param name="path">the path of the eml file.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		/// <returns>A value indicating success or not.</returns>
		public static bool EmlToPst(string path, string pstPath)
		{
			bool result = false;

			Log.Info("Checking file: " + path);

			if (Directory.Exists(path))
			{
				EmlDirectoryToPst(path, pstPath);
				result = true;
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
				string message = string.Format(
					CultureInfo.InvariantCulture,
					"Duplicate key mapping! Folder Id: {1} Name: {0} ",
					dbxFolder.FolderName,
					dbxFolder.FolderId);

				Log.Info(message);
			}
			else
			{
				mappings.Add(dbxFolder.FolderId, pstFolder.EntryID);
			}
		}

		private static MAPIFolder CopyChildFolderToPst(
			IDictionary<uint, string> mappings,
			PstOutlook pstOutlook,
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

			pstFolder = PstOutlook.AddFolderSafe(
				parentFolder, dbxFolder.FolderName);

			Marshal.ReleaseComObject(parentFolder);

			return pstFolder;
		}

		private static void CopyEmlToPst(
			PstOutlook pstOutlook, MAPIFolder pstFolder, string emlFile)
		{
			if (!string.IsNullOrWhiteSpace(emlFile) && File.Exists(emlFile))
			{
				string msgFile = GetTemporaryMsgFile();

				try
				{
					Converter.ConvertEmlToMsg(emlFile, msgFile);
				}
				catch (System.Exception exception) when
					(exception is InvalidCastException ||
					exception is NullReferenceException)
				{
					Log.Error(exception.ToString());
				}

				pstOutlook.AddMsgFile(pstFolder, msgFile);

				File.Delete(msgFile);
			}
		}

		private static void CopyMessages(
			PstOutlook pstOutlook, MAPIFolder pstFolder, DbxFolder dbxFolder)
		{
			// for each message
			DbxMessage dbxMessage;

			do
			{
				dbxMessage = dbxFolder.GetNextMessage();

				CopyMessageToPst(pstOutlook, pstFolder, dbxMessage);
			}
			while (dbxMessage != null);

			Marshal.ReleaseComObject(pstFolder);
		}

		private static void CopyMessageToPst(
			PstOutlook pstOutlook, MAPIFolder pstFolder, DbxMessage dbxMessage)
		{
			if (dbxMessage != null && dbxMessage.Message.Length > 0)
			{
				// Need to get the rfc email as a stream, then
				// convert the stream to a MSG file, import the
				// MSG file into the Pst, finally move the message
				using Stream emailStream = dbxMessage.MessageStream;

				string msgFile = GetTemporaryMsgFile();

				using Stream msgStream =
					PstOutlook.GetMsgFileStream(msgFile);

				try
				{
					Converter.ConvertEmlToMsg(emailStream, msgStream);
				}
				catch (System.Exception exception) when
					(exception is InvalidCastException ||
					exception is NullReferenceException)
				{
					Log.Error(exception.ToString());
				}

				msgStream.Dispose();

				pstOutlook.AddMsgFile(pstFolder, msgFile);

				File.Delete(msgFile);
			}
		}

		/// <summary>
		/// Dbx directory to pst.
		/// </summary>
		/// <param name="emlFilesPath">The path to dbx folders to
		/// migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		private static void EmlDirectoryToPst(
			string emlFilesPath, string pstPath)
		{
			PstOutlook pstOutlook = new ();
			Store pstStore = pstOutlook.CreateStore(pstPath);

			if (pstStore == null)
			{
				Log.Error("PST store not created");
			}
			else
			{
				string baseName = Path.GetFileNameWithoutExtension(pstPath);

				IEnumerable<string> emlFiles =
					EmlMessages.GetFiles(emlFilesPath);

				if (emlFiles.Any())
				{
					MAPIFolder pstFolder =
						PstOutlook.GetTopLevelFolder(pstStore, baseName);

					foreach (string file in emlFiles)
					{
						CopyEmlToPst(pstOutlook, pstFolder, file);
					}

					Marshal.ReleaseComObject(pstFolder);
				}
			}
		}

		/// <summary>
		/// Dbx files to pst.
		/// </summary>
		/// <param name="filePath">The file path to migrate.</param>
		/// <param name="pstPath">The path to pst file to copy to.</param>
		private static void EmlFileToPst(string filePath, string pstPath)
		{
			PstOutlook pstOutlook = new ();
			Store pstStore = pstOutlook.CreateStore(pstPath);

			string baseName = Path.GetFileNameWithoutExtension(pstPath);

			MAPIFolder pstFolder =
				PstOutlook.GetTopLevelFolder(pstStore, baseName);

			CopyEmlToPst(pstOutlook, pstFolder, filePath);

			Marshal.ReleaseComObject(pstFolder);
		}

		private static string GetTemporaryMsgFile()
		{
			string msgFile = Path.GetTempFileName();

			// A 0 byte sized file is created.  Need to remove it.
			File.Delete(msgFile);
			msgFile = Path.ChangeExtension(msgFile, ".msg");

			return msgFile;
		}
	}
}
