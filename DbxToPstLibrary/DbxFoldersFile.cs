/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxFoldersFile.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Globalization;
using System.IO;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx folders files class.
	/// </summary>
	public class DbxFoldersFile : DbxFile
	{
		/// <summary>
		/// The id index of the folder.
		/// </summary>
		public const int Id = 0x00;
		/// <summary>
		/// The parent id index of the folder.
		/// </summary>
		public const int ParentId = 0x01;
		/// <summary>
		/// The name index of the folder.
		/// </summary>
		public const int Name = 0x02;
		/// <summary>
		/// The flags index of the folder.
		/// </summary>
		public const int Flags = 0x06;

		private const int TreeNodeSize = 0x27c;

		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private DbxTree tree;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="DbxFoldersFile"/> class.
		/// </summary>
		/// <param name="filePath">The path of the dbx file.</param>
		public DbxFoldersFile(string filePath)
			: base(filePath)
		{
		}

		/// <summary>
		/// List folders method.
		/// </summary>
		public void List()
		{
			if (tree != null)
			{
				byte[] fileBytes = GetFileBytes();

				foreach (uint index in tree.FolderInformationIndexes)
				{
					DbxFolderIndexedItem item = new (fileBytes, index);
					item.ReadIndex(fileBytes, index);

					DbxFolderIndex folderIndex = item.FolderIndex;

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						Id,
						folderIndex.FolderId);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						ParentId,
						folderIndex.FolderParentId);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						Name,
						folderIndex.FolderName);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						Name,
						folderIndex.FolderFileName);
					Log.Info(message);
				}
			}
		}

		/// <summary>
		/// Migrate folder method.
		/// </summary>
		/// <param name="folderName">The path of the dbx folder file.</param>
		public static void MigrateFolder(string folderName)
		{
			if (!string.IsNullOrWhiteSpace(folderName))
			{
				string foldersPath = Path.GetDirectoryName(folderName);
				string filePath = Path.Combine(foldersPath, folderName);

				bool exists = File.Exists(filePath);

				if (exists == false)
				{
					Log.Warn(
						filePath + " specified in Folders.dbx not present");
				}
				else
				{
					DbxMessagesFile messagesFile = new (filePath);

					DbxFileType check = messagesFile.Header.FileType;

					if (check != DbxFileType.MessageFile)
					{
						Log.Error(filePath + " not actually a messagess file");

						// Attempt to process the individual files.
					}
					else
					{
						messagesFile.ReadTree();
					}
				}
			}
		}

		/// <summary>
		/// Migrate folders method.
		/// </summary>
		public void MigrateFolders()
		{
			if (tree != null)
			{
				byte[] fileBytes = GetFileBytes();

				foreach (uint index in tree.FolderInformationIndexes)
				{
					DbxFolderIndexedItem item = new(fileBytes, index);
					item.ReadIndex(fileBytes, index);

					DbxFolderIndex folderIndex = item.FolderIndex;

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						Name,
						folderIndex.FolderName);
					Log.Info(message);
				}
			}
		}

		/// <summary>
		/// Read the tree method.
		/// </summary>
		public void ReadTree()
		{
			byte[] fileBytes = GetFileBytes();

			tree = new (fileBytes, Header.MainTreeAddress, Header.FolderCount);
		}
	}
}
