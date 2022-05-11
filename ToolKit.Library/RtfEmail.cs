/////////////////////////////////////////////////////////////////////////////
// <copyright file="RtfEmail.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Text.RegularExpressions;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Represents a RTF email body.
	/// </summary>
	public static class RtfEmail
	{
		/// <summary>
		/// Trim the end of the RTF body.
		/// </summary>
		/// <param name="rtfBody">The RTF body to trim.</param>
		/// <returns>The trimmed RTF body.</returns>
		public static byte[] Trim(byte[] rtfBody)
		{
			if (rtfBody != null)
			{
				byte[] footer = new byte[10];
				int offset = rtfBody.Length - footer.Length;
				Array.Copy(rtfBody, offset, footer, 0, footer.Length);

				byte[] checkBytes = new byte[]
				{
				92, 112, 97, 114, 13, 10, 125, 13, 10, 0
				};

				bool confirm = CheckBytes(footer, footer.Length, 0);

				if (confirm == true)
				{
					int counts = 0;
					int removeCount = 0;
					byte[] endLine = new byte[6];
					Array.Copy(footer, endLine, 6);

					while (confirm == true)
					{
						int off = rtfBody.Length - footer.Length -
							removeCount - endLine.Length;
						confirm = CheckBytes(rtfBody, endLine.Length, off);

						if (confirm == true)
						{
							counts++;
							removeCount = endLine.Length * counts;
						}
					}

					// Re-attach optmized footer.
					int size = rtfBody.Length - removeCount;
					byte[] newBody = new byte[size];
					int begin = rtfBody.Length - footer.Length - removeCount;
					Array.Copy(rtfBody, newBody, begin);
					Array.Copy(footer, 0, newBody, begin, footer.Length);

					rtfBody = newBody;
				}
			}

			return rtfBody;
		}

		private static bool CheckBytes(
			byte[] bytesToCheck, int count, int offset)
		{
			bool confirm = true;

			byte[] checkBytes = new byte[]
			{
				92, 112, 97, 114, 13, 10, 125, 13, 10, 0
			};

			for (int index = 0; index < count; index++)
			{
				int subOffset = offset + index;
				byte byteToCheck = bytesToCheck[subOffset];
				byte checkByte = checkBytes[index];

				if (byteToCheck != checkByte)
				{
					confirm = false;
					break;
				}
			}

			return confirm;
		}
	}
}
