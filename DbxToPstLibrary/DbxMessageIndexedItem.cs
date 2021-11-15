/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxMessageIndexedItem.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx message indexed item.
	/// </summary>
	public class DbxMessageIndexedItem : DbxIndexedItem
	{
		/// <summary>
		/// The id index of the folder.
		/// </summary>
		public const int Id = 0x00;
		/// <summary>
		/// The name index of the folder.
		/// </summary>
		public const int Name = 0x02;

		private DbxMessageIndex messageIndex;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="DbxMessageIndexedItem"/> class.
		/// </summary>
		/// <param name="fileBytes">The bytes of the file.</param>
		/// <param name="address">The address of the item with in
		/// the file.</param>
		public DbxMessageIndexedItem(byte[] fileBytes, uint address)
			: base(fileBytes, address)
		{
			messageIndex = new DbxMessageIndex();
		}

		public DbxMessageIndex MessageIndex { get { return messageIndex; } }
	}
}
