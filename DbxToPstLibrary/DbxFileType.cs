/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxFileType.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx file type.
	/// </summary>
	public enum DbxFileType
	{
		/// <summary>
		/// Unknown file type.
		/// </summary>
		Unknown,

		/// <summary>
		/// Message file.
		/// </summary>
		MessageFile,

		/// <summary>
		/// Folder file.
		/// </summary>
		FolderFile,

		/// <summary>
		/// Off line file.
		/// </summary>
		OffLine,

		/// <summary>
		/// Pop3 file.
		/// </summary>
		Pop3uidl
	}
}
