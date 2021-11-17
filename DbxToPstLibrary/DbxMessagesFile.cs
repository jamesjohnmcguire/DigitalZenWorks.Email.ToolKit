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
		/// List messages method.
		/// </summary>
		public void List()
		{
			if (Tree != null)
			{
				byte[] fileBytes = GetFileBytes();

				foreach (uint index in Tree.FolderInformationIndexes)
				{
					DbxMessageIndexedItem item = new (fileBytes);
					item.ReadIndex(index);

					DbxMessageIndex messageIndex = item.MessageIndex;

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						DbxMessageIndexedItem.SenderName,
						messageIndex.SenderName);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						DbxMessageIndexedItem.SenderEmailAddress,
						messageIndex.SenderEmailAddress);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						DbxMessageIndexedItem.ReceivedTime,
						messageIndex.ReceivedTime);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						DbxMessageIndexedItem.Subject,
						messageIndex.Subject);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						DbxMessageIndexedItem.ReceiptentName,
						messageIndex.ReceiptentName);
					Log.Info(message);

					message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						DbxMessageIndexedItem.ReceiptentEmailAddress,
						messageIndex.ReceiptentEmailAddress);
					Log.Info(message);
				}
			}
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
					DbxMessageIndexedItem item = new (fileBytes);
					item.ReadIndex(index);

					DbxMessageIndex messageIndex = item.MessageIndex;

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"item value[{0}] is {1}",
						"some",
						messageIndex.Id);
					Log.Info(message);
				}
			}
		}
	}
}
