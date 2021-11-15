/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxMessagesFile.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Globalization;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx emails file.
	/// </summary>
	public class DbxMessagesFile : DbxFile
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="DbxMessagesFile"/> class.
		/// </summary>
		/// <param name="filePath">The path of the dbx file.</param>
		public DbxMessagesFile(string filePath)
			: base(filePath)
		{
		}

		/// <summary>
		/// Migrate messages method.
		/// </summary>
		public void MigrateMessages()
		{
			if (Tree != null)
			{
				byte[] fileBytes = GetFileBytes();

				foreach (uint index in Tree.FolderInformationIndexes)
				{
					DbxMessageIndexedItem item = new (fileBytes, index);
					item.ReadIndex(fileBytes, index);

					DbxMessageIndex messageIndex = item.MessageIndex;

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						"some",
						messageIndex.MessageId);
					Log.Info(message);
				}
			}
		}
	}
}
