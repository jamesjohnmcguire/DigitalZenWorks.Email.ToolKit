/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxIndexedItem.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx indexed item class.
	/// </summary>
	public class DbxIndexedItem
	{
		// Somewhat arbitrary, as other references have this as 0x20, but other
		// notes indicate this may not enough.
		private const int MaximumIndexes = 0x40;

		private byte[] bodyBytes;
		private uint[] indexes;

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxIndexedItem"/>
		/// class.
		/// </summary>
		/// <param name="fileBytes">The bytes of the file.</param>
		/// <param name="address">The address of the item with in
		/// the file.</param>
		public DbxIndexedItem(byte[] fileBytes, uint address)
		{
			indexes = new uint[MaximumIndexes];

			byte[] initialBytes = new byte[12];

			Array.Copy(fileBytes, address, initialBytes, 0, 12);

			// It will be easier to work with integers as opposed to bytes.
			uint[] initialArray = Bytes.ToIntegerArray(initialBytes);

			if (initialArray[0] != address)
			{
				throw new DbxException("Wrong object marker!");
			}

			uint bodyLength = initialArray[1];
			ushort indexLength = Bytes.ToShort(initialBytes, 8);
			byte itemsCount = initialBytes[10];
			byte itemsChangedCount = initialBytes[11];

			uint offset = address + 12;

			bodyBytes = new byte[bodyLength];
			Array.Copy(fileBytes, offset, bodyBytes, 0, bodyLength);

			// why is this needed? what the rationale?
			var shifter = itemsCount << 2;

			uint pointer = bodyBytes[shifter];

			int itemsCountBytes = itemsCount * 4;

			for (uint index = 0; index < itemsCountBytes; index += 4)
			{
				byte rawValue = bodyBytes[index];
				bool isDirect = Bytes.GetBit(rawValue, 7);

				byte index2 = (byte)(rawValue & 0x7F);
				// uint value = Bytes.ToIntegerLimit(bodyBytes, index + 1, 3);
				uint value = index;

				if (isDirect == true)
				{
					value++;
					SetIndex(index2, value);
				}
				else
				{
					value = bodyBytes[index + 1];
					offset = 16;
					value = offset + value;
					SetIndex(index2, value);
				}
			}
		}

		private void SetIndex(uint index, uint value)
		{
			indexes[index] = value;
		}
	}
}
