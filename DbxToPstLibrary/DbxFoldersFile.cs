/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxFoldersFile.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Globalization;

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

		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private const int TreeNodeSize = 0x27c;
		private DbxTree tree;

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxFoldersFile"/> class.
		/// </summary>
		/// <param name="filePath">The path of the dbx file.</param>
		public DbxFoldersFile(string filePath)
			: base(filePath)
		{
		}

		public void List()
		{
			if (tree != null)
			{
				byte[] fileBytes = GetFileBytes();

				foreach (uint index in tree.FolderInformationIndexes)
				{
					DbxIndexedItem item = new(fileBytes, index);

					uint value = item.GetValue(Id);

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						Id,
						value);
					Log.Info(message);

					value = item.GetValue(ParentId);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						ParentId,
						value);
					Log.Info(message);

					string name = item.GetString(Name);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						Name,
						name);
					Log.Info(message);
					break;
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
