/////////////////////////////////////////////////////////////////////////////
// <copyright file="MapiItemComparer.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ToolKit.Library
{
	/// <summary>
	/// Provides comparision support for Outlook MAPI items.
	/// </summary>
	public static class MapiItemComparer
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// Gets the item's hash.
		/// </summary>
		/// <param name="item">The items to compute.</param>
		/// <returns>The item's hash.</returns>
		public static byte[] GetItemHash(MailItem item)
		{
			return null;
		}

		private static byte[] GetBody(MailItem mailItem)
		{
			Encoding encoding = Encoding.UTF8;
			byte[] body = encoding.GetBytes(mailItem.Body);
			byte[] htmlBody = encoding.GetBytes(mailItem.HTMLBody);
			byte[] rtfBody = encoding.GetBytes(mailItem.RTFBody);

			long bufferSize =
				body.LongLength + htmlBody.LongLength + rtfBody.LongLength;
			byte[] allBody = new byte[bufferSize];

			// combine the parts
			Array.Copy(body, allBody, body.Length);

			Array.Copy(
				htmlBody, 0, allBody, body.LongLength, htmlBody.LongLength);

			long destinationIndex = body.LongLength + htmlBody.LongLength;
			Array.Copy(
				rtfBody, 0, allBody, destinationIndex, rtfBody.LongLength);

			return allBody;
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage(
			"StyleCop.CSharp.NamingRules",
			"SA1305:Field names should not use Hungarian notation",
			Justification = "It isn't hungarian notation.")]
		private static string GetRecipients(MailItem mailItem)
		{
			string recipients = string.Empty;
			List<string> toList = new ();
			List<string> ccList = new ();
			List<string> bccList = new ();

			foreach (Recipient recipient in mailItem.Recipients)
			{
				string formattedRecipient = string.Format(
					CultureInfo.InvariantCulture,
					"{0} <{1}>; ",
					recipient.Name,
					recipient.Address);

				OlMailRecipientType recipientType =
					(OlMailRecipientType)recipient.Type;
				switch (recipientType)
				{
					case OlMailRecipientType.olTo:
						toList.Add(formattedRecipient);
						break;
					case OlMailRecipientType.olCC:
						ccList.Add(formattedRecipient);
						break;
					case OlMailRecipientType.olBCC:
						bccList.Add(formattedRecipient);
						break;
					case OlMailRecipientType.olOriginator:
						Log.Warn("Ignoring olOriginator recipient type");
						break;
					default:
						Log.Warn("Ignoring uknown recipient type");
						break;
				}
			}

			toList.Sort();
			ccList.Sort();
			bccList.Sort();

			foreach (string formattedRecipient in toList)
			{
				recipients += formattedRecipient;
			}

			foreach (string formattedRecipient in ccList)
			{
				recipients += formattedRecipient;
			}

			foreach (string formattedRecipient in bccList)
			{
				recipients += formattedRecipient;
			}

			return recipients;
		}
	}
}
