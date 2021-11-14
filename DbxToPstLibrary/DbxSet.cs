/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxSet.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx set class.
	/// </summary>
	public class DbxSet
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private readonly DbxFoldersFile foldersFile;

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxSet"/> class.
		/// </summary>
		/// <param name="path">The path of the dbx set.</param>
		public DbxSet(string path)
		{
			string extension = Path.GetExtension(path);

			if (string.IsNullOrEmpty(extension))
			{
				path = Path.Combine(path, "Folders.dbx");
			}

			bool exists = File.Exists(path);

			if (exists == false)
			{
				Log.Error( path + " not present");

				// Attempt to process the individual files.
			}
			else
			{
				foldersFile = new (path);

				if (foldersFile.Header.FileType != DbxFileType.FolderFile)
				{
					Log.Error("Folders.dbx not actually folders file");

					// Attempt to process the individual files.
				}
				else
				{
					foldersFile.ReadTree();
				}
			}
		}

		/// <summary>
		/// List method.
		/// </summary>
		public void List()
		{
			foldersFile.List();
		}
	}
}
