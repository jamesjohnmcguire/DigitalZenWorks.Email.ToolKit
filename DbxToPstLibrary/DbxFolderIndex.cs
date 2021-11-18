/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxFolderIndex.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx folder index class.
	/// </summary>
	public class DbxFolderIndex
	{
		/// <summary>
		/// Gets or sets the folder file name.
		/// </summary>
		/// <value>The folder file name.</value>
		public string FolderFileName { get; set; }

		/// <summary>
		/// Gets or sets the folder id.
		/// </summary>
		/// <value>The folder id.</value>
		public uint FolderId { get; set; }

		/// <summary>
		/// Gets or sets the folder name.
		/// </summary>
		/// <value>The folder name.</value>
		public string FolderName { get; set; }

		/// <summary>
		/// Gets or sets the folder parent id.
		/// </summary>
		/// <value>The folder parent id.</value>
		public uint FolderParentId { get; set; }
	}
}
