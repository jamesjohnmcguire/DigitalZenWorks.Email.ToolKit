/////////////////////////////////////////////////////////////////////////////
// <copyright file="Migrate.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Email.DbxOutlookExpress;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

[assembly: CLSCompliant(false)]

namespace DbxToPstLibrary
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
			DbxSet dbxSet = new (dbxFoldersPath);
			DbxFolder dbxFolder;

			PstOutlook pstOutlook = new ();
			Store pstStore = pstOutlook.CreateStore(pstPath);

			MAPIFolder rootFolder = pstStore.GetRootFolder();

			IDictionary<uint, string> mappings = new Dictionary<uint, string>();

			do
			{
				dbxFolder = dbxSet.GetNextFolder();

				if (dbxFolder != null)
				{
					// add folder to pst
					if (dbxFolder.FolderParentId == 0)
					{
						// top level folder
						MAPIFolder pstFolder =
							rootFolder.Folders.Add(dbxFolder.FolderName);

						mappings.Add(dbxFolder.FolderId, pstFolder.EntryID);
					}
					else
					{
						// need to figure out parent in pst
						bool keyExists =
							mappings.ContainsKey(dbxFolder.FolderParentId);

						if (keyExists == false)
						{
							Log.Warn("Parent key not found in mappings: " +
								dbxFolder.FolderParentId);
						}
						else
						{
							string entryId = mappings[dbxFolder.FolderParentId];
							MAPIFolder parentFolder =
								pstOutlook.GetFolderFromID(entryId, pstStore);

							MAPIFolder pstFolder =
								parentFolder.Folders.Add(dbxFolder.FolderName);

							mappings.Add(dbxFolder.FolderId, pstFolder.EntryID);
						}
					}

					// for each message
					DbxMessage dbxMessage;

					do
					{
						dbxMessage = dbxFolder.GetNextMessage();

						if (dbxMessage != null)
						{
							// add message to pst
						}
					}
					while (dbxMessage != null);
				}
			}
			while (dbxFolder != null);
		}

		/// <summary>
		/// Dbx files to pst.
		/// </summary>
		/// <param name="filePath">The file path to migrate.</param>
		public static void DbxFileToPst(string filePath)
		{
			DbxSet dbxSet = new (filePath);
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

			if (Directory.Exists(path))
			{
				DbxDirectoryToPst(path, pstPath);
				result = true;
			}
			else if (File.Exists(path))
			{
				DbxFileToPst(path);
				result = true;
			}
			else
			{
				Log.Error("Invalid path");
			}

			return result;
		}

		/// <summary>
		/// Transfers a folder from a data source to a destination source.
		/// </summary>
		/// <param name="id">The folder id.</param>
		public static void TransferFolder(int id)
		{
		}
	}
}
