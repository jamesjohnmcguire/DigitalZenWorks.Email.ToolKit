/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxFoldersFile.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx folders files class.
	/// </summary>
	public class DbxFoldersFile : DbxFile
	{
		private const int TreeNodeSize = 0x27c;

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxFoldersFile"/> class.
		/// </summary>
		/// <param name="filePath">The path of the dbx file.</param>
		public DbxFoldersFile(string filePath)
			: base(filePath)
		{
		}

		public void ReadTree()
		{
			byte[] treeBytes = new byte[TreeNodeSize];
			Array.Copy(
				FileBytes, Header.MainTreeAddress, treeBytes, 0, TreeNodeSize);

			DbxTree tree = new DbxTree(
				treeBytes, Header.MainTreeAddress, Header.FolderCount);
		}
	}
}
