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
using System.IO;
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

		private static byte[] GetAttachments(MailItem mailItem)
		{
			byte[] attachments = null;

			try
			{
				string basePath = Path.GetTempPath();

				foreach (Attachment attachment in mailItem.Attachments)
				{
					Encoding encoding = Encoding.UTF8;

					int intType = (int)attachment.Type;

					string index = attachment.Index.ToString(
						CultureInfo.InvariantCulture);
					string position = attachment.Position.ToString(
						CultureInfo.InvariantCulture);
					string attachmentType =
						intType.ToString(CultureInfo.InvariantCulture);

					string metaData = string.Format(
						CultureInfo.InvariantCulture,
						"{0}{1}{2}{3}{4}{5}",
						attachment.DisplayName,
						attachment.FileName,
						index,
						attachment.PathName,
						position,
						attachmentType);

					byte[] metaDataBytes = encoding.GetBytes(metaData);

					if (attachments == null)
					{
						attachments = metaDataBytes;
					}
					else
					{
						attachments =
							MergeByteArrays(attachments, metaDataBytes);
					}

					string filePath = basePath + attachment.FileName;
					attachment.SaveAsFile(filePath);

					byte[] fileBytes = File.ReadAllBytes(filePath);

					attachments = MergeByteArrays(attachments, fileBytes);
				}
			}
			catch (System.Exception exception) when
			(exception is ArgumentException ||
			exception is ArgumentNullException ||
			exception is ArgumentOutOfRangeException ||
			exception is ArrayTypeMismatchException ||
			exception is InvalidCastException ||
			exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return attachments;
		}

		private static byte[] GetBody(MailItem mailItem)
		{
			byte[] allBody = null;

			try
			{
				Encoding encoding = Encoding.UTF8;
				byte[] body = encoding.GetBytes(mailItem.Body);
				byte[] htmlBody = encoding.GetBytes(mailItem.HTMLBody);
				byte[] rtfBody = mailItem.RTFBody as byte[];

				allBody = MergeByteArrays(body, htmlBody);
				allBody = MergeByteArrays(allBody, rtfBody);
			}
			catch (System.Exception exception) when
			(exception is ArgumentException ||
			exception is ArgumentNullException ||
			exception is ArgumentOutOfRangeException ||
			exception is ArrayTypeMismatchException ||
			exception is InvalidCastException ||
			exception is RankException)
			{
				Log.Error(exception.ToString());
			}

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

		private static byte[] GetSimpleProperties(MailItem mailItem)
		{
			string data1 = string.Format(
				CultureInfo.InvariantCulture,
				"{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}",
				mailItem.BCC,
				mailItem.BillingInformation,
				mailItem.Categories,
				mailItem.CC,
				mailItem.Companies,
				mailItem.ConversationID,
				mailItem.ConversationIndex,
				mailItem.ConversationTopic,
				mailItem.FlagRequest,
				mailItem.MessageClass,
				mailItem.Mileage,
				mailItem.OutlookVersion,
				mailItem.PermissionTemplateGuid,
				mailItem.ReceivedByEntryID,
				mailItem.ReceivedByName,
				mailItem.ReceivedOnBehalfOfEntryID);

			string data2 = string.Format(
				CultureInfo.InvariantCulture,
				"{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}",
				mailItem.ReceivedOnBehalfOfName,
				mailItem.ReminderSoundFile,
				mailItem.ReplyRecipientNames,
				mailItem.RetentionPolicyName,
				mailItem.SenderEmailAddress,
				mailItem.SenderEmailType,
				mailItem.SenderName,
				mailItem.SentOnBehalfOfName,
				mailItem.Subject,
				mailItem.TaskSubject,
				mailItem.To,
				mailItem.VotingOptions,
				mailItem.VotingResponse);

			string data = string.Format(
				CultureInfo.InvariantCulture, "{0}{1}", data1, data2);

			return null;
		}

		private static byte[] MergeByteArrays(byte[] buffer1, byte[] buffer2)
		{
			byte[] newBuffer = null;

			try
			{
				Encoding encoding = Encoding.UTF8;

				long bufferSize =
					buffer1.LongLength + buffer2.LongLength;
				newBuffer = new byte[bufferSize];

				// combine the parts
				Array.Copy(buffer1, newBuffer, buffer1.LongLength);

				Array.Copy(
					buffer2,
					0,
					newBuffer,
					buffer1.LongLength,
					buffer2.LongLength);
			}
			catch (System.Exception exception) when
			(exception is ArgumentException ||
			exception is ArgumentNullException ||
			exception is ArgumentOutOfRangeException ||
			exception is ArrayTypeMismatchException ||
			exception is InvalidCastException ||
			exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return newBuffer;
		}
	}
}
