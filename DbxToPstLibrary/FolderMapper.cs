/////////////////////////////////////////////////////////////////////////////
// <copyright file="FolderMapper.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Folder mapper class.
	/// </summary>
	public class FolderMapper
	{
		/// <summary>
		/// Gets or sets the dbx folder id.
		/// </summary>
		/// <value>The dbx folder id.</value>
		public int DbxFolderId { get; set; }

		/// <summary>
		/// Gets or sets the MAPI entry id.
		/// </summary>
		/// <value>The MAPI entry id.</value>
		public string EntryId { get; set; }
	}
}
