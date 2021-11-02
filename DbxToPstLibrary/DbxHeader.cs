/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxHeader.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx header.
	/// </summary>
	public class DbxHeader
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private int fileInfoLength;
		private DbxFileType fileType;
		private int lastSegmentAddress;

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxHeader"/> class.
		/// </summary>
		/// <param name="headerBytes">An array of bytes representing the
		/// file header.</param>
		public DbxHeader(byte[] headerBytes)
		{
			byte[] checkBytes = new byte[] { 0xCF, 0xAD, 0x12, 0xFE, 0xC5,
				0xFD, 0x74, 0x6F, 0x66, 0xE3, 0xD1, 0x11, 0x9A, 0x4E, 0x00,
				0xC0, 0x4F, 0xA3, 0x09, 0xD4, 0x05, 0x00, 0x00, 0x00, 0x05,
				0x00, 0x00, 0x00 };

			if (headerBytes != null)
			{
				fileType = GetFileType(headerBytes);

				for (int index = 0; index < checkBytes.Length; index++)
				{
					if (index == 4)
					{
						continue;
					}

					ConfirmByte(headerBytes, index, checkBytes[index]);
				}

				fileInfoLength = BytesToInteger(headerBytes, 0x1C);
				lastSegmentAddress = BytesToInteger(headerBytes, 0x24);
			}
		}

		private static int BytesToInteger(byte[] bytes, int index)
		{
			int result;
			byte[] testBytes = new byte[4];
			Array.Copy(bytes, index, testBytes, 0, 4);

			// Dbx files are apprentely stored as little endian.
			if (BitConverter.IsLittleEndian == false)
			{
				Array.Reverse(testBytes);
			}

			result = BitConverter.ToInt32(testBytes, 0);

			return result;
		}

		private static bool ConfirmByte(
			byte[] bytes, int index, byte checkValue)
		{
			bool confirm = false;

			byte byteToCheck = bytes[index];

			if (byteToCheck == checkValue)
			{
				confirm = true;
			}
			else
			{
				Log.Warn("bytes not matching at" +
					index.ToString(CultureInfo.InvariantCulture));
			}

			return confirm;
		}

		private static DbxFileType GetFileType(byte[] bytes)
		{
			DbxFileType fileType = DbxFileType.Unknown;
			string message;
			byte byteToCheck = bytes[4];

			switch (byteToCheck)
			{
				case 0xC5:
					fileType = DbxFileType.MessageFile;
					break;
				case 0xC6:
					fileType = DbxFileType.FolderFile;
					break;
				case 0xC7:
					fileType = DbxFileType.Pop3uidl;
					break;
				case 0x30:
					if (bytes[5] == 0x9D && bytes[6] == 0xFE &&
						bytes[7] == 0x26)
					{
						fileType = DbxFileType.OffLine;
					}
					else
					{
						message = string.Format(
							CultureInfo.InvariantCulture,
							"File type unknown {0} {1} {2} {3}",
							0x30,
							bytes[5],
							bytes[6],
							bytes[7]);

						Log.Warn(message);
					}

					break;
				default:
					message = string.Format(
						CultureInfo.InvariantCulture,
						"File type unknown {0} {1} {2} {3}",
						bytes[4],
						bytes[5],
						bytes[6],
						bytes[7]);

					Log.Warn(message);
					break;
			}

			return fileType;
		}
	}
}
