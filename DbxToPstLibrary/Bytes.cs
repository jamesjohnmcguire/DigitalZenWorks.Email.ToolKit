/////////////////////////////////////////////////////////////////////////////
// <copyright file="Bytes.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Bytes class.
	/// </summary>
	public static class Bytes
	{
		/// <summary>
		/// Get bit value, as boolean.
		/// </summary>
		/// <param name="rawValue">The byte with the bit to check.</param>
		/// <param name="bitNumber">The 0 based index of the bit
		/// to retrieve.</param>
		/// <returns>The bit value, as boolean.</returns>
		public static bool GetBit(byte rawValue, byte bitNumber)
		{
			bool bitValue = false;

			// 0 based
			byte shifter2 = (byte)(1 << bitNumber);
			byte bit = (byte)(rawValue & shifter2);

			if (bit != 0)
			{
				bitValue = true;
			}

			return bitValue;
		}

		/// <summary>
		/// To integer method.
		/// </summary>
		/// <param name="bytes">The source bytes.</param>
		/// <param name="index">The index with in the bytes to copy.</param>
		/// <returns>An integer of the bytes values.</returns>
		public static uint ToInteger(byte[] bytes, int index)
		{
			uint result = ToIntegerLimit(bytes, index, 4);

			return result;
		}

		/// <summary>
		/// To integer limit method.
		/// </summary>
		/// <param name="bytes">The source bytes.</param>
		/// <param name="index">The index with in the bytes to copy.</param>
		/// <param name="limit">The amount of bytes to copy.</param>
		/// <returns>An integer of the bytes values.</returns>
		public static uint ToIntegerLimit(byte[] bytes, int index, int limit)
		{
			uint result;
			byte[] testBytes = new byte[limit];
			Array.Copy(bytes, index, testBytes, 0, limit);

			// Dbx files are apprentely stored as little endian.
			if (BitConverter.IsLittleEndian == false)
			{
				Array.Reverse(testBytes);
			}

			result = BitConverter.ToUInt32(testBytes, 0);

			return result;
		}

		/// <summary>
		/// To integer array method.
		/// </summary>
		/// <param name="bytes">The source bytes.</param>
		/// <returns>An integer array of the bytes values.</returns>
		public static uint[] ToIntegerArray(byte[] bytes)
		{
			uint[] integerArray = null;

			if (bytes != null)
			{
				int size = bytes.Length / sizeof(uint);
				integerArray = new uint[size];
				Buffer.BlockCopy(bytes, 0, integerArray, 0, bytes.Length);
			}

			return integerArray;
		}

		/// <summary>
		/// To short method.
		/// </summary>
		/// <param name="bytes">The source bytes.</param>
		/// <param name="index">The index with in the bytes to copy.</param>
		/// <returns>An integer of the bytes values.</returns>
		public static ushort ToShort(byte[] bytes, int index)
		{
			ushort result;
			byte[] testBytes = new byte[2];
			Array.Copy(bytes, index, testBytes, 0, 2);

			// Dbx files are apprentely stored as little endian.
			if (BitConverter.IsLittleEndian == false)
			{
				Array.Reverse(testBytes);
			}

			result = BitConverter.ToUInt16(testBytes, 0);

			return result;
		}
	}
}
