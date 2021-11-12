/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxIndexedItem.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Text;

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

		private readonly byte[] bodyBytes;
		private readonly uint[] indexes;

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
			byte itemsCount = initialBytes[10];

			uint offset = address + 12;

			bodyBytes = new byte[bodyLength];
			Array.Copy(fileBytes, offset, bodyBytes, 0, bodyLength);

			uint itemsCountBytes = (uint)itemsCount * 4;

			for (uint index = 0; index < itemsCountBytes; index += 4)
			{
				byte rawValue = bodyBytes[index];
				bool isDirect = Bytes.GetBit(rawValue, 7);

				byte index2 = (byte)(rawValue & 0x7F);
				uint value = index;

				if (isDirect == true)
				{
					value++;
					SetIndex(index2, value);
				}
				else
				{
					value = bodyBytes[index + 1];
					offset = itemsCountBytes;
					value = offset + value;
					SetIndex(index2, value);
				}
			}
		}

		/// <summary>
		/// Get a string value from the indexed item.
		/// </summary>
		/// <param name="index">The index item to retrieve.</param>
		/// <returns>The value of the itemed item.</returns>
		public string GetString(uint index)
		{
			string item = null;
			uint subIndex = indexes[index];

			if (subIndex > 0)
			{
				uint end = subIndex;
				byte check;

				do
				{
					check = bodyBytes[end];

					if (check == 0)
					{
						break;
					}

					end++;
				}
				while (check > 0);

				int length = (int)(end - subIndex);

				item =
					Encoding.ASCII.GetString(bodyBytes, (int)subIndex, length);
			}

			return item;
		}

		/// <summary>
		/// Get the values from the indexed item.
		/// </summary>
		/// <param name="index">The index item to retrieve.</param>
		/// <returns>The value of the itemed item.</returns>
		public uint GetValue(uint index)
		{
			uint item = 0;
			uint subIndex = indexes[index];

			if (subIndex > 0)
			{
				item = bodyBytes[subIndex];
				item = Bytes.ToIntegerLimit(bodyBytes, subIndex, 3);
			}

			return item;
		}

		private void SetIndex(uint index, uint value)
		{
			indexes[index] = value;
		}
	}
}
