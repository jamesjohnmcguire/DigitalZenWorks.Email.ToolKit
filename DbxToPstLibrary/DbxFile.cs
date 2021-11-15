/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxFile.cs" company="James John McGuire">
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
	/// Dbx file class.
	/// </summary>
	public class DbxFile
	{
		private byte[] fileBytes;
		private string folderPath;

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxFile"/> class.
		/// </summary>
		/// <param name="filePath">The path of the dbx file.</param>
		public DbxFile(string filePath)
		{
			folderPath = filePath;

			fileBytes = File.ReadAllBytes(filePath);

			byte[] headerBytes = new byte[0x24bc];
			Array.Copy(fileBytes, headerBytes, 0x24bc);

			Header = new (headerBytes);
		}

		/// <summary>
		/// Gets or sets the dbx file header.
		/// </summary>
		/// <value>The dbx file header.</value>
		public DbxHeader Header { get; set; }

		/// <summary>
		/// Gets the dbx folder file path.
		/// </summary>
		/// <value>The dbx folder file path.</value>
		public string FolderPath { get { return folderPath; } }

		/// <summary>
		/// Gets the file bytes.
		/// </summary>
		/// <returns>The file bytes.</returns>
		public byte[] GetFileBytes()
		{
			return fileBytes;
		}
	}
}
