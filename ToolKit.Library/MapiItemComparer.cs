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

		private static byte[] CopyIntToByteArray(
			byte[] bytes, int index, int value)
		{
			byte byteValue1 = (byte)value;
			byte byteValue2 = (byte)(value >> 8);
			byte byteValue3 = (byte)(value >> 0x10);
			byte byteValue4 = (byte)(value >> 0x18);

			bytes[index] = byteValue1;
			index++;
			bytes[index] = byteValue2;
			index++;
			bytes[index] = byteValue3;
			index++;
			bytes[index] = byteValue4;

			return bytes;
		}

		private static byte[] GetActions(MailItem mailItem)
		{
			byte[] actions = null;

			try
			{
				foreach (Microsoft.Office.Interop.Outlook.Action action in
					mailItem.Actions)
				{
					Encoding encoding = Encoding.UTF8;

					int copyLikeEnum = (int)action.CopyLike;
					int enabledBool = Convert.ToInt32(action.Enabled);
					int replyStyleEnum = (int)action.ReplyStyle;
					int responseStyleEnum = (int)action.ResponseStyle;
					int showOnEnum = (int)action.ShowOn;

					string copyLike =
						copyLikeEnum.ToString(CultureInfo.InvariantCulture);
					string enabled =
						enabledBool.ToString(CultureInfo.InvariantCulture);
					string replyStyle =
						replyStyleEnum.ToString(CultureInfo.InvariantCulture);
					string responseStyle = responseStyleEnum.ToString(
						CultureInfo.InvariantCulture);
					string showOn =
						showOnEnum.ToString(CultureInfo.InvariantCulture);

					string metaData = string.Format(
						CultureInfo.InvariantCulture,
						"{0}{1}{2}{3}{4}{5}{6}",
						copyLike,
						enabled,
						action.Name,
						action.Prefix,
						replyStyle,
						responseStyle,
						showOn);

					byte[] metaDataBytes = encoding.GetBytes(metaData);

					if (actions == null)
					{
						actions = metaDataBytes;
					}
					else
					{
						actions = MergeByteArrays(actions, metaDataBytes);
					}
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

			return actions;
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

		private static ushort GetBooleans(MailItem mailItem)
		{
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

			return boolHolder;
		}

		private static string GetDateTimes(MailItem mailItem)
		{
			string deferredDeliveryTime = mailItem.DeferredDeliveryTime.ToString("O");
			string expiryTime = mailItem.ExpiryTime.ToString("O");
			string receivedTime = mailItem.ReceivedTime.ToString("O");
			string reminderTime = mailItem.ReminderTime.ToString("O");
			string retentionExpirationDate =
				mailItem.RetentionExpirationDate.ToString("O");
			string sentOn = mailItem.SentOn.ToString("O");
			string taskCompletedDate = mailItem.TaskCompletedDate.ToString("O");
			string taskDueDate = mailItem.TaskDueDate.ToString("O");
			string taskStartDate = mailItem.TaskStartDate.ToString("O");

			string data = string.Format(
				CultureInfo.InvariantCulture,
				"{0}{1}{2}{3}{4}{5}{6}{7}{8}",
				deferredDeliveryTime,
				expiryTime,
				receivedTime,
				reminderTime,
				retentionExpirationDate,
				sentOn,
				taskCompletedDate,
				taskDueDate,
				taskStartDate);

			return data;
		}

		private static byte[] GetEnums(MailItem mailItem)
		{
			int bodyFormat = (int)mailItem.BodyFormat;
			int itemClass = (int)mailItem.Class;
			int importance = (int)mailItem.Importance;
			int markForDownload = (int)mailItem.MarkForDownload;
			int permission = (int)mailItem.Permission;
			int permissionService = (int)mailItem.PermissionService;
			int sensitivity = (int)mailItem.Sensitivity;

			string internetCodepage = mailItem.InternetCodepage.ToString(
				CultureInfo.InvariantCulture);
			string size = mailItem.Size.ToString(CultureInfo.InvariantCulture);

			// 9 ints * size of int
			int bufferSize = 9 * 4;
			byte[] buffer = new byte[bufferSize];

			int index = 0;
			buffer = CopyIntToByteArray(buffer, index, bodyFormat);
			index += 4;
			buffer = CopyIntToByteArray(buffer, index, itemClass);
			index += 4;
			buffer = CopyIntToByteArray(buffer, index, importance);
			index += 4;
			buffer = CopyIntToByteArray(buffer, index, markForDownload);
			index += 4;
			buffer = CopyIntToByteArray(buffer, index, permission);
			index += 4;
			buffer = CopyIntToByteArray(buffer, index, permissionService);
			index += 4;
			buffer = CopyIntToByteArray(buffer, index, sensitivity);
			index += 4;
			buffer =
				CopyIntToByteArray(buffer, index, mailItem.InternetCodepage);
			index += 4;
			buffer = CopyIntToByteArray(buffer, index, mailItem.Size);

			return buffer;
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
			string header = mailItem.PropertyAccessor.GetProperty(
				"http://schemas.microsoft.com/mapi/proptag/0x007D001F");

			string data1 = string.Format(
				CultureInfo.InvariantCulture,
				"{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}",
				mailItem.BCC,
				mailItem.BillingInformation,
				mailItem.Body,
				mailItem.Categories,
				mailItem.CC,
				mailItem.Companies,
				mailItem.ConversationID,
				mailItem.ConversationIndex,
				mailItem.ConversationTopic,
				mailItem.FlagRequest,
				header,
				mailItem.HTMLBody,
				mailItem.MessageClass,
				mailItem.Mileage,
				mailItem.OutlookVersion,
				mailItem.PermissionTemplateGuid);

			string data2 = string.Format(
				CultureInfo.InvariantCulture,
				"{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}",
				mailItem.ReceivedByEntryID,
				mailItem.ReceivedByName,
				mailItem.ReceivedOnBehalfOfEntryID,
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

			string deferredDeliveryTime = mailItem.DeferredDeliveryTime.ToString("O");
			string expiryTime = mailItem.ExpiryTime.ToString("O");
			string receivedTime = mailItem.ReceivedTime.ToString("O");
			string reminderTime = mailItem.ReminderTime.ToString("O");
			string retentionExpirationDate =
				mailItem.RetentionExpirationDate.ToString("O");
			string sentOn = mailItem.SentOn.ToString("O");
			string taskCompletedDate = mailItem.TaskCompletedDate.ToString("O");
			string taskDueDate = mailItem.TaskDueDate.ToString("O");
			string taskStartDate = mailItem.TaskStartDate.ToString("O");

			string data = string.Format(
				CultureInfo.InvariantCulture, "{0}{1}", data1, data2);

			return null;
		}

		private static byte[] GetUserProperties(MailItem mailItem)
		{
			byte[] properties = null;

			try
			{
				foreach (UserProperty property in mailItem.UserProperties)
				{
					Encoding encoding = Encoding.UTF8;

					int typeEnum = (int)property.Type;

					string typeValue =
						typeEnum.ToString(CultureInfo.InvariantCulture);
					string value =
						property.Value.ToString(CultureInfo.InvariantCulture);

					string metaData = string.Format(
						CultureInfo.InvariantCulture,
						"{0}{1}{2}{3}{4}{5}",
						property.Formula,
						property.Name,
						typeValue,
						property.ValidationFormula,
						property.ValidationText,
						value);

					byte[] metaDataBytes = encoding.GetBytes(metaData);

					if (properties == null)
					{
						properties = metaDataBytes;
					}
					else
					{
						properties =
							MergeByteArrays(properties, metaDataBytes);
					}
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

			return properties;
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
