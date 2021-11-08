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
				foreach (uint index in tree.FolderInformationIndexes)
				{

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
