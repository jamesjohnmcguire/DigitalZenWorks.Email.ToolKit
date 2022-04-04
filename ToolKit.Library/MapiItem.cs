/////////////////////////////////////////////////////////////////////////////
// <copyright file="MapiItem.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Provides comparision support for Outlook MAPI items.
	/// </summary>
	public static class MapiItem
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// Deletes the given item.
		/// </summary>
		/// <param name="item">The item to delete.</param>
		public static void DeleteItem(object item)
		{
			switch (item)
			{
				case AppointmentItem appointmentItem:
					appointmentItem.Delete();
					Marshal.ReleaseComObject(appointmentItem);
					break;
				case ContactItem contactItem:
					contactItem.Delete();
					Marshal.ReleaseComObject(contactItem);
					break;
				case DistListItem distListItem:
					distListItem.Delete();
					Marshal.ReleaseComObject(distListItem);
					break;
				case DocumentItem documentItem:
					documentItem.Delete();
					Marshal.ReleaseComObject(documentItem);
					break;
				case JournalItem journalItem:
					journalItem.Delete();
					Marshal.ReleaseComObject(journalItem);
					break;
				case MailItem mailItem:
					mailItem.Delete();
					Marshal.ReleaseComObject(mailItem);
					break;
				case MeetingItem meetingItem:
					meetingItem.Delete();
					Marshal.ReleaseComObject(meetingItem);
					break;
				case NoteItem noteItem:
					noteItem.Delete();
					Marshal.ReleaseComObject(noteItem);
					break;
				case PostItem postItem:
					postItem.Delete();
					Marshal.ReleaseComObject(postItem);
					break;
				case RemoteItem remoteItem:
					remoteItem.Delete();
					Marshal.ReleaseComObject(remoteItem);
					break;
				case ReportItem reportItem:
					reportItem.Delete();
					Marshal.ReleaseComObject(reportItem);
					break;
				case TaskItem taskItem:
					taskItem.Delete();
					Marshal.ReleaseComObject(taskItem);
					break;
				case TaskRequestAcceptItem taskRequestAcceptItem:
					taskRequestAcceptItem.Delete();
					Marshal.ReleaseComObject(taskRequestAcceptItem);
					break;
				case TaskRequestDeclineItem taskRequestDeclineItem:
					taskRequestDeclineItem.Delete();
					Marshal.ReleaseComObject(taskRequestDeclineItem);
					break;
				case TaskRequestItem taskRequestItem:
					taskRequestItem.Delete();
					Marshal.ReleaseComObject(taskRequestItem);
					break;
				case TaskRequestUpdateItem taskRequestUpdateItem:
					taskRequestUpdateItem.Delete();
					Marshal.ReleaseComObject(taskRequestUpdateItem);
					break;
				default:
					Log.Warn(
						"folder item of unknown type: " + item.ToString());
					break;
			}

			Marshal.ReleaseComObject(item);
		}

		/// <summary>
		/// Gets the item's hash.
		/// </summary>
		/// <param name="path">The path of the curent folder.</param>
		/// <param name="mailItem">The items to compute.</param>
		/// <returns>The item's hash encoded in base 64.</returns>
		public static string GetItemHash(string path, MailItem mailItem)
		{
			string hashBase64 = null;
			try
			{
				if (mailItem != null)
				{
					byte[] finalBuffer = GetItemBytes(path, mailItem);

					using SHA256 hasher = SHA256.Create();

					byte[] hashValue = hasher.ComputeHash(finalBuffer);
					hashBase64 = Convert.ToBase64String(hashValue);
				}
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is InvalidCastException ||
				exception is OutOfMemoryException ||
				exception is RankException)
			{
				LogException(path, string.Empty, mailItem);
				Log.Error(exception.ToString());
			}

			return hashBase64;
		}

		private static long ArrayCopyConditional(
			ref byte[] finalBuffer, long currentIndex, byte[] nextBuffer)
		{
			if (nextBuffer != null)
			{
				Array.Copy(
					nextBuffer,
					0,
					finalBuffer,
					currentIndex,
					nextBuffer.LongLength);
				currentIndex += nextBuffer.LongLength;
			}

			return currentIndex;
		}

		private static byte[] CopyIntToByteArray(
			byte[] bytes, long index, int value)
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

		private static byte[] CopyUshortToByteArray(
			byte[] bytes, long index, ushort value)
		{
			byte byteValue1 = (byte)value;
			byte byteValue2 = (byte)(value >> 8);

			bytes[index] = byteValue1;
			index++;
			bytes[index] = byteValue2;

			return bytes;
		}

		private static byte[] GetActions(string path, MailItem mailItem)
		{
			byte[] actions = null;

			try
			{
				int total = mailItem.Actions.Count;

				for (int index = 1; index <= total; index++)
				{
					Microsoft.Office.Interop.Outlook.Action action =
						mailItem.Actions[index];

					Encoding encoding = Encoding.UTF8;

					int copyLikeEnum = (int)action.CopyLike;
					bool enabledBool = action.Enabled;
					int enabledInt = Convert.ToInt32(enabledBool);
					int replyStyleEnum = (int)action.ReplyStyle;
					int responseStyleEnum = (int)action.ResponseStyle;
					int showOnEnum = (int)action.ShowOn;

					string copyLike =
						copyLikeEnum.ToString(CultureInfo.InvariantCulture);
					string enabled =
						enabledInt.ToString(CultureInfo.InvariantCulture);
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

					Marshal.ReleaseComObject(action);
				}
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is System.Runtime.InteropServices.COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return actions;
		}

		private static byte[] GetAttachments(string path, MailItem mailItem)
		{
			byte[] attachments = null;

			try
			{
				string basePath = Path.GetTempPath();

				int total = mailItem.Attachments.Count;

				for (int index = 1; index <= total; index++)
				{
					Attachment attachment =
						mailItem.Attachments[index];

					Encoding encoding = Encoding.UTF8;

					int attachmentIndex = attachment.Index;
					string indexValue = attachmentIndex.ToString(
						CultureInfo.InvariantCulture);

					int positionValue = attachment.Position;
					string position = positionValue.ToString(
						CultureInfo.InvariantCulture);

					int intType = (int)attachment.Type;
					string attachmentType =
						intType.ToString(CultureInfo.InvariantCulture);

					string metaData = string.Format(
						CultureInfo.InvariantCulture,
						"{0}{1}{2}{3}{4}",
						attachment.DisplayName,
						attachment.FileName,
						indexValue,
						position,
						attachmentType);

					try
					{
						metaData += attachment.PathName;
					}
					catch (System.Runtime.InteropServices.COMException)
					{
					}

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

					Marshal.ReleaseComObject(attachment);
				}
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is System.Runtime.InteropServices.COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return attachments;
		}

		private static byte[] GetBody(string path, MailItem mailItem)
		{
			byte[] allBody = null;

			try
			{
				Encoding encoding = Encoding.UTF8;

				string body = mailItem.Body;
				string htmlBody = mailItem.HTMLBody;

				byte[] bodyBytes = encoding.GetBytes(body);
				byte[] htmlBodyBytes = encoding.GetBytes(htmlBody);
				byte[] rtfBody = mailItem.RTFBody as byte[];

				allBody = MergeByteArrays(bodyBytes, htmlBodyBytes);
				allBody = MergeByteArrays(allBody, rtfBody);
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is System.Runtime.InteropServices.COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return allBody;
		}

		private static ushort GetBooleans(string path, MailItem mailItem)
		{
			ushort boolHolder = 0;

			try
			{
				bool rawValue = mailItem.AlternateRecipientAllowed;
				boolHolder = SetBit(boolHolder, 0, rawValue);

				rawValue = mailItem.AutoForwarded;
				boolHolder = SetBit(boolHolder, 1, rawValue);

				rawValue = mailItem.AutoResolvedWinner;
				boolHolder = SetBit(boolHolder, 2, rawValue);

				rawValue = mailItem.DeleteAfterSubmit;
				boolHolder = SetBit(boolHolder, 3, rawValue);

				rawValue = mailItem.IsMarkedAsTask;
				boolHolder = SetBit(boolHolder, 4, rawValue);

				rawValue = mailItem.NoAging;
				boolHolder = SetBit(boolHolder, 5, rawValue);

				rawValue = mailItem.OriginatorDeliveryReportRequested;
				boolHolder = SetBit(boolHolder, 6, rawValue);

				rawValue = mailItem.ReadReceiptRequested;
				boolHolder = SetBit(boolHolder, 7, rawValue);

				rawValue = mailItem.RecipientReassignmentProhibited;
				boolHolder = SetBit(boolHolder, 8, rawValue);

				rawValue = mailItem.ReminderOverrideDefault;
				boolHolder = SetBit(boolHolder, 9, rawValue);

				rawValue = mailItem.ReminderPlaySound;
				boolHolder = SetBit(boolHolder, 10, rawValue);

				rawValue = mailItem.ReminderSet;
				boolHolder = SetBit(boolHolder, 11, rawValue);

				rawValue = mailItem.Saved;
				boolHolder = SetBit(boolHolder, 12, rawValue);

				rawValue = mailItem.Sent;
				boolHolder = SetBit(boolHolder, 13, rawValue);

				rawValue = mailItem.Submitted;
				boolHolder = SetBit(boolHolder, 14, rawValue);

				rawValue = mailItem.UnRead;
				boolHolder = SetBit(boolHolder, 15, rawValue);
			}
			catch (System.Runtime.InteropServices.COMException exception)
			{
				Log.Error(exception.ToString());
			}

			return boolHolder;
		}

		private static long GetBufferSize(
			byte[] actions,
			byte[] attachments,
			byte[] dateTimes,
			byte[] enums,
			byte[] rtfBody,
			byte[] strings,
			byte[] userProperties)
		{
			long bufferSize = 0;

			if (actions != null)
			{
				bufferSize += actions.LongLength;
			}

			if (attachments != null)
			{
				bufferSize += attachments.LongLength;
			}

			if (dateTimes != null)
			{
				bufferSize += dateTimes.LongLength;
			}

			if (enums != null)
			{
				bufferSize += enums.LongLength;
			}

			if (rtfBody != null)
			{
				bufferSize += rtfBody.LongLength;
			}

			if (strings != null)
			{
				bufferSize += strings.LongLength;
			}

			if (userProperties != null)
			{
				bufferSize += userProperties.LongLength;
			}

			bufferSize += 2;

			return bufferSize;
		}

		private static byte[] GetDateTimes(string path, MailItem mailItem)
		{
			byte[] data = null;

			try
			{
				DateTime deferredDeliveryTimeDateTime =
					mailItem.DeferredDeliveryTime;
				DateTime expiryTimeDateTime = mailItem.ExpiryTime;
				DateTime receivedTimeDateTime = mailItem.ReceivedTime;
				DateTime reminderTimeDateTime = mailItem.ReminderTime;
				DateTime retentionExpirationDateDateTime =
					mailItem.RetentionExpirationDate;
				DateTime sentOnDateTime = mailItem.SentOn;
				DateTime taskCompletedDateDateTime =
					mailItem.TaskCompletedDate;
				DateTime taskDueDateDateTime = mailItem.TaskDueDate;
				DateTime taskStartDateDateTime = mailItem.TaskStartDate;

				string deferredDeliveryTime =
					deferredDeliveryTimeDateTime.ToString("O");
				string expiryTime = expiryTimeDateTime.ToString("O");
				string receivedTime = receivedTimeDateTime.ToString("O");
				string reminderTime = reminderTimeDateTime.ToString("O");
				string retentionExpirationDate =
					retentionExpirationDateDateTime.ToString("O");
				string sentOn = sentOnDateTime.ToString("O");
				string taskCompletedDate =
					taskCompletedDateDateTime.ToString("O");
				string taskDueDate = taskDueDateDateTime.ToString("O");
				string taskStartDate = taskStartDateDateTime.ToString("O");

				string buffer = string.Format(
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

				Encoding encoding = Encoding.UTF8;
				data = encoding.GetBytes(buffer);
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is System.Runtime.InteropServices.COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return data;
		}

		private static byte[] GetEnums(string path, MailItem mailItem)
		{
			byte[] buffer = null;

			try
			{
				int bodyFormat = (int)mailItem.BodyFormat;
				int itemClass = (int)mailItem.Class;
				int importance = (int)mailItem.Importance;
				int markForDownload = (int)mailItem.MarkForDownload;
				int permission = 0;
				int permissionService = (int)mailItem.PermissionService;
				int sensitivity = (int)mailItem.Sensitivity;

				try
				{
					permission = (int)mailItem.Permission;
				}
				catch (System.Runtime.InteropServices.COMException)
				{
				}

				// 9 ints * size of int
				int bufferSize = 9 * 4;
				buffer = new byte[bufferSize];

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
				buffer = CopyIntToByteArray(
					buffer, index, mailItem.InternetCodepage);
				index += 4;
				buffer = CopyIntToByteArray(buffer, index, mailItem.Size);
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is System.Runtime.InteropServices.COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				string sentOn = mailItem.SentOn.ToString(
					"yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

				string message = string.Format(
					CultureInfo.InvariantCulture,
					"Item: {0}: From: {1}: {2} Subject: {3}",
					sentOn,
					mailItem.SenderName,
					mailItem.SenderEmailAddress,
					mailItem.Subject);

				Log.Error("Exception at: " + path);
				Log.Error(message);
				Log.Error(exception.ToString());
			}

			return buffer;
		}

		private static byte[] GetItemBytes(string path, MailItem mailItem)
		{
			byte[] finalBuffer = null;

			try
			{
				if (mailItem != null)
				{
					ushort booleans = GetBooleans(path, mailItem);

					byte[] actions = GetActions(path, mailItem);
					byte[] attachments = GetAttachments(path, mailItem);
					byte[] dateTimes = GetDateTimes(path, mailItem);
					byte[] enums = GetEnums(path, mailItem);
					byte[] rtfBody = null;

					try
					{
						rtfBody = mailItem.RTFBody as byte[];
					}
					catch (System.Runtime.InteropServices.COMException)
					{
						string sentOn = mailItem.SentOn.ToString(
							"yyyy-MM-dd HH:mm:ss",
							CultureInfo.InvariantCulture);

						string message = string.Format(
							CultureInfo.InvariantCulture,
							"Item: {0}: From: {1}: {2} Subject: {3}",
							sentOn,
							mailItem.SenderName,
							mailItem.SenderEmailAddress,
							mailItem.Subject);

						Log.Error("Exception on RTFBody at: " + path);
						Log.Error(message);
					}

					byte[] strings = GetStringProperties(path, mailItem);
					byte[] userProperties = GetUserProperties(path, mailItem);

					long bufferSize = GetBufferSize(
						actions,
						attachments,
						dateTimes,
						enums,
						rtfBody,
						strings,
						userProperties);

					finalBuffer = new byte[bufferSize];

					// combine the parts
					long currentIndex = ArrayCopyConditional(
						ref finalBuffer, 0, actions);

					currentIndex = ArrayCopyConditional(
						ref finalBuffer, currentIndex, attachments);
					currentIndex = ArrayCopyConditional(
						ref finalBuffer, currentIndex, dateTimes);
					currentIndex = ArrayCopyConditional(
						ref finalBuffer, currentIndex, enums);
					currentIndex = ArrayCopyConditional(
						ref finalBuffer, currentIndex, rtfBody);
					currentIndex = ArrayCopyConditional(
						ref finalBuffer, currentIndex, strings);
					currentIndex = ArrayCopyConditional(
						ref finalBuffer, currentIndex, userProperties);

					finalBuffer = CopyUshortToByteArray(
						finalBuffer, currentIndex, booleans);
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

			return finalBuffer;
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage(
			"StyleCop.CSharp.NamingRules",
			"SA1305:Field names should not use Hungarian notation",
			Justification = "It isn't hungarian notation.")]
		private static string GetRecipients(string path, MailItem mailItem)
		{
			string recipients = string.Empty;
			List<string> toList = new ();
			List<string> ccList = new ();
			List<string> bccList = new ();

			int total = mailItem.Recipients.Count;

			for (int index = 1; index <= total; index++)
			{
				Recipient recipient = mailItem.Recipients[index];
				string name = recipient.Name;
				string address = recipient.Address;

				string formattedRecipient = string.Format(
					CultureInfo.InvariantCulture,
					"{0} <{1}>; ",
					name,
					address);

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

				Marshal.ReleaseComObject(recipient);
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

		private static byte[] GetStringProperties(
			string path, MailItem mailItem)
		{
			byte[] data = null;

			try
			{
				string bcc = mailItem.BCC;
				string billingInformation = mailItem.BillingInformation;
				string body = mailItem.Body;
				string categories = mailItem.Categories;
				string cc = mailItem.CC;
				string companies = mailItem.Companies;
				string conversationID = mailItem.ConversationID;
				string conversationTopic = mailItem.ConversationTopic;
				string flagRequest = mailItem.FlagRequest;
				string header = mailItem.PropertyAccessor.GetProperty(
					"http://schemas.microsoft.com/mapi/proptag/0x007D001F");
				string htmlBody = mailItem.HTMLBody;
				string messageClass = mailItem.MessageClass;
				string mileage = mailItem.Mileage;
				string receivedByEntryID = mailItem.ReceivedByEntryID;
				string receivedByName = mailItem.ReceivedByName;
				string receivedOnBehalfOfEntryID =
					mailItem.ReceivedOnBehalfOfEntryID;
				string receivedOnBehalfOfName =
					mailItem.ReceivedOnBehalfOfName;
				string reminderSoundFile = mailItem.ReminderSoundFile;
				string replyRecipientNames = mailItem.ReplyRecipientNames;
				string retentionPolicyName = mailItem.RetentionPolicyName;
				string senderEmailAddress = mailItem.SenderEmailAddress;
				string senderEmailType = mailItem.SenderEmailType;
				string senderName = mailItem.SenderName;
				string sentOnBehalfOfName = mailItem.SentOnBehalfOfName;
				string subject = mailItem.Subject;
				string taskSubject = mailItem.TaskSubject;
				string to = mailItem.To;
				string votingOptions = mailItem.VotingOptions;
				string votingResponse = mailItem.VotingResponse;

				StringBuilder builder = new ();
				builder.Append(bcc);
				builder.Append(billingInformation);
				builder.Append(body);
				builder.Append(categories);
				builder.Append(cc);
				builder.Append(companies);
				builder.Append(conversationID);
				builder.Append(conversationTopic);
				builder.Append(flagRequest);
				builder.Append(header);
				builder.Append(htmlBody);
				builder.Append(messageClass);
				builder.Append(mileage);
				builder.Append(receivedByEntryID);
				builder.Append(receivedByName);
				builder.Append(receivedOnBehalfOfEntryID);
				builder.Append(receivedOnBehalfOfName);
				builder.Append(reminderSoundFile);
				builder.Append(replyRecipientNames);
				builder.Append(retentionPolicyName);
				builder.Append(senderEmailAddress);
				builder.Append(senderEmailAddress);
				builder.Append(senderName);
				builder.Append(sentOnBehalfOfName);
				builder.Append(subject);
				builder.Append(taskSubject);
				builder.Append(to);
				builder.Append(votingOptions);
				builder.Append(votingResponse);

				string buffer = builder.ToString();

				Encoding encoding = Encoding.UTF8;
				data = encoding.GetBytes(buffer);
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is System.Runtime.InteropServices.COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return data;
		}

		private static byte[] GetUserProperties(string path, MailItem mailItem)
		{
			byte[] properties = null;

			try
			{
				int total = mailItem.UserProperties.Count;

				for (int index = 1; index <= total; index++)
				{
					UserProperty property = mailItem.UserProperties[index];

					Encoding encoding = Encoding.UTF8;

					int typeEnum = (int)property.Type;
					var propertyValue = property.Value;

					string typeValue =
						typeEnum.ToString(CultureInfo.InvariantCulture);
					string value =
						propertyValue.ToString(CultureInfo.InvariantCulture);

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

					Marshal.ReleaseComObject(property);
				}
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is System.Runtime.InteropServices.COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return properties;
		}

		private static void LogException(
			string path, string extraInformation, MailItem mailItem)
		{
			string sentOn = mailItem.SentOn.ToString(
				"yyyy-MM-dd HH:mm:ss",
				CultureInfo.InvariantCulture);

			string message = string.Format(
				CultureInfo.InvariantCulture,
				"Item: {0}: From: {1}: {2} Subject: {3}",
				sentOn,
				mailItem.SenderName,
				mailItem.SenderEmailAddress,
				mailItem.Subject);

			Log.Error("Exception " + extraInformation + "at: " + path);
			Log.Error(message);
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
