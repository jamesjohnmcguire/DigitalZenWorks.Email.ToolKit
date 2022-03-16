﻿/////////////////////////////////////////////////////////////////////////////
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

			int bodyFormat = (int)mailItem.BodyFormat;
			int itemClass = (int)mailItem.Class;
			int downloadState = (int)mailItem.DownloadState;
			int flagStatus = (int)mailItem.FlagStatus;
			int importance = (int)mailItem.Importance;
			int markForDownload = (int)mailItem.MarkForDownload;
			int permission = (int)mailItem.Permission;
			int permissionService = (int)mailItem.PermissionService;
			int remoteStatus = (int)mailItem.RemoteStatus;
			int sensitivity = (int)mailItem.Sensitivity;

			string internetCodepage = mailItem.InternetCodepage.ToString(
				CultureInfo.InvariantCulture);
			string size = mailItem.Size.ToString(CultureInfo.InvariantCulture);

			string data3 = string.Format(
				CultureInfo.InvariantCulture,
				"{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}",
				bodyFormat.ToString(CultureInfo.InvariantCulture),
				itemClass.ToString(CultureInfo.InvariantCulture),
				downloadState.ToString(CultureInfo.InvariantCulture),
				flagStatus.ToString(CultureInfo.InvariantCulture),
				importance.ToString(CultureInfo.InvariantCulture),
				markForDownload.ToString(CultureInfo.InvariantCulture),
				permission.ToString(CultureInfo.InvariantCulture),
				permissionService.ToString(CultureInfo.InvariantCulture),
				remoteStatus.ToString(CultureInfo.InvariantCulture),
				sensitivity.ToString(CultureInfo.InvariantCulture),
				internetCodepage,
				size);

			ushort boolHolder = 0;
			boolHolder = SetBit(
				boolHolder, 0, mailItem.AlternateRecipientAllowed);
			boolHolder = SetBit(boolHolder, 1, mailItem.AutoForwarded);
			boolHolder = SetBit(boolHolder, 2, mailItem.AutoResolvedWinner);
			boolHolder = SetBit(boolHolder, 3, mailItem.DeleteAfterSubmit);
			boolHolder = SetBit(boolHolder, 4, mailItem.IsMarkedAsTask);
			boolHolder = SetBit(boolHolder, 5, mailItem.NoAging);
			boolHolder = SetBit(
				boolHolder, 6, mailItem.OriginatorDeliveryReportRequested);
			boolHolder = SetBit(boolHolder, 7, mailItem.ReadReceiptRequested);
			boolHolder = SetBit(
				boolHolder, 8, mailItem.RecipientReassignmentProhibited);
			boolHolder = SetBit(
				boolHolder, 9, mailItem.ReminderOverrideDefault);
			boolHolder = SetBit(boolHolder, 10, mailItem.ReminderPlaySound);
			boolHolder = SetBit(boolHolder, 11, mailItem.ReminderSet);
			boolHolder = SetBit(boolHolder, 12, mailItem.Saved);
			boolHolder = SetBit(boolHolder, 13, mailItem.Sent);
			boolHolder = SetBit(boolHolder, 14, mailItem.Submitted);
			boolHolder = SetBit(boolHolder, 15, mailItem.UnRead);

			/*
Actions
Conflicts
CreationTime
DeferredDeliveryTime
ExpiryTime
FormDescription
LastModificationTime
Links
ReceivedTime
ReminderTime
RetentionExpirationDate
SaveSentMessageFolder
SentOn
TaskCompletedDate
TaskDueDate
	TaskStartDate
ToDoTaskOrdinal
UserProperties
*/

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

		private static byte SetBit(byte holder, byte bitIndex, bool value)
		{
			int intValue = Convert.ToInt32(value);

			// 0 based
			int shifter = intValue << bitIndex;
			int intHolder = holder | shifter;
			holder = (byte)intHolder;

			return holder;
		}

		private static ushort SetBit(ushort holder, byte bitIndex, bool value)
		{
			int intValue = Convert.ToInt32(value);

			// 0 based
			int shifter = intValue << bitIndex;
			int intHolder = holder | shifter;
			holder = (ushort)intHolder;

			return holder;
		}
	}
}
