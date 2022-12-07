/////////////////////////////////////////////////////////////////////////////
// <copyright file="MapiItem.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Common.Utilities;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

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
			try
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
						string message = "Folder item of unknown type";
						if (item != null)
						{
							message += ": " + item.ToString();
						}

						Log.Warn(message);
						break;
				}

				Marshal.ReleaseComObject(item);
			}
			catch (COMException exception)
			{
				Log.Error(exception.ToString());
			}
		}

		/// <summary>
		/// Gets the item's hash.
		/// </summary>
		/// <param name="path">The path of the current folder.</param>
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

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <param name="mailItem">The MailItem to check.</param>
		/// <returns>The synoses of the item.</returns>
		public static string GetItemSynopses(MailItem mailItem)
		{
			string synopses = null;

			if (mailItem != null)
			{
				string sentOn = mailItem.SentOn.ToString(
					"yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

				synopses = string.Format(
					CultureInfo.InvariantCulture,
					"{0}: From: {1}: {2} Subject: {3}",
					sentOn,
					mailItem.SenderName,
					mailItem.SenderEmailAddress,
					mailItem.Subject);
			}

			return synopses;
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="item">The item to move.</param>
		/// <param name="destination">The destination folder.</param>
		public static void Moveitem(object item, MAPIFolder destination)
		{
			try
			{
				switch (item)
				{
					case AppointmentItem appointmentItem:
						appointmentItem.Move(destination);
						Marshal.ReleaseComObject(appointmentItem);
						break;
					case ContactItem contactItem:
						contactItem.Move(destination);
						Marshal.ReleaseComObject(contactItem);
						break;
					case DistListItem distListItem:
						distListItem.Move(destination);
						Marshal.ReleaseComObject(distListItem);
						break;
					case DocumentItem documentItem:
						documentItem.Move(destination);
						Marshal.ReleaseComObject(documentItem);
						break;
					case JournalItem journalItem:
						journalItem.Move(destination);
						Marshal.ReleaseComObject(journalItem);
						break;
					case MailItem mailItem:
						mailItem.Move(destination);
						Marshal.ReleaseComObject(mailItem);
						break;
					case MeetingItem meetingItem:
						meetingItem.Move(destination);
						Marshal.ReleaseComObject(meetingItem);
						break;
					case NoteItem noteItem:
						noteItem.Move(destination);
						Marshal.ReleaseComObject(noteItem);
						break;
					case PostItem postItem:
						postItem.Move(destination);
						Marshal.ReleaseComObject(postItem);
						break;
					case RemoteItem remoteItem:
						remoteItem.Move(destination);
						Marshal.ReleaseComObject(remoteItem);
						break;
					case ReportItem reportItem:
						reportItem.Move(destination);
						Marshal.ReleaseComObject(reportItem);
						break;
					case TaskItem taskItem:
						taskItem.Move(destination);
						Marshal.ReleaseComObject(taskItem);
						break;
					case TaskRequestAcceptItem taskRequestAcceptItem:
						taskRequestAcceptItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestAcceptItem);
						break;
					case TaskRequestDeclineItem taskRequestDeclineItem:
						taskRequestDeclineItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestDeclineItem);
						break;
					case TaskRequestItem taskRequestItem:
						taskRequestItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestItem);
						break;
					case TaskRequestUpdateItem taskRequestUpdateItem:
						taskRequestUpdateItem.Move(destination);
						Marshal.ReleaseComObject(taskRequestUpdateItem);
						break;
					default:
						string message = "Folder item of unknown type";
						if (item != null)
						{
							message += ": " + item.ToString();
						}

						Log.Warn(message);
						break;
				}

				Marshal.ReleaseComObject(item);
			}
			catch (COMException exception)
			{
				Log.Error(exception.ToString());
			}
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="item">The item to move.</param>
		/// <param name="destination">The destination folder.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public static async Task MoveitemAsync(
			object item, MAPIFolder destination)
		{
			CancellationTokenSource source = new CancellationTokenSource();

			try
			{
				source.CancelAfter(TimeSpan.FromSeconds(5));

				switch (item)
				{
					case AppointmentItem appointmentItem:
						await Task.Run(() =>
							appointmentItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(appointmentItem);
						break;
					case ContactItem contactItem:
						await Task.Run(() =>
							contactItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(contactItem);
						break;
					case DistListItem distListItem:
						await Task.Run(() =>
							distListItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(distListItem);
						break;
					case DocumentItem documentItem:
						await Task.Run(() =>
							documentItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(documentItem);
						break;
					case JournalItem journalItem:
						await Task.Run(() =>
							journalItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(journalItem);
						break;
					case MailItem mailItem:
						await Task.Run(() =>
							mailItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(mailItem);
						break;
					case MeetingItem meetingItem:
						await Task.Run(() =>
							meetingItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(meetingItem);
						break;
					case NoteItem noteItem:
						await Task.Run(() =>
							noteItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(noteItem);
						break;
					case PostItem postItem:
						await Task.Run(() =>
							postItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(postItem);
						break;
					case RemoteItem remoteItem:
						await Task.Run(() =>
							remoteItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(remoteItem);
						break;
					case ReportItem reportItem:
						await Task.Run(() =>
							reportItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(reportItem);
						break;
					case TaskItem taskItem:
						await Task.Run(() =>
							taskItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(taskItem);
						break;
					case TaskRequestAcceptItem taskRequestAcceptItem:
						await Task.Run(() =>
							taskRequestAcceptItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(taskRequestAcceptItem);
						break;
					case TaskRequestDeclineItem taskRequestDeclineItem:
						await Task.Run(() =>
							taskRequestDeclineItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(taskRequestDeclineItem);
						break;
					case TaskRequestItem taskRequestItem:
						await Task.Run(() =>
							taskRequestItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(taskRequestItem);
						break;
					case TaskRequestUpdateItem taskRequestUpdateItem:
						await Task.Run(() =>
							taskRequestUpdateItem.Move(destination)).
								ConfigureAwait(false);
						Marshal.ReleaseComObject(taskRequestUpdateItem);
						break;
					default:
						string message = "Folder item of unknown type";
						if (item != null)
						{
							message += ": " + item.ToString();
						}

						Log.Warn(message);
						break;
				}

				Marshal.ReleaseComObject(item);
			}
			catch (System.Exception exception) when
				(exception is COMException ||
				exception is OperationCanceledException)
			{
				Log.Error(exception.ToString());
			}
			finally
			{
				source.Dispose();
			}
		}

		/// <summary>
		/// Removes the MimeOLE version number.
		/// </summary>
		/// <param name="header">The header to check.</param>
		/// <returns>The modified header.</returns>
		public static string RemoveMimeOleVersion(string header)
		{
			string pattern = @"(?<=Produced By Microsoft MimeOLE)" +
				@" V(\d+)\.(\d+)\.(\d+)\.(\d+)";

			header = Regex.Replace(
				header, pattern, string.Empty, RegexOptions.ExplicitCapture);

			return header;
		}

		private static byte[] GetActions(MailItem mailItem)
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
						actions = BitBytes.MergeByteArrays(actions, metaDataBytes);
					}

					Marshal.ReleaseComObject(action);
				}
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Warn(exception.ToString());
			}

			return actions;
		}

		private static byte[] GetAttachments(MailItem mailItem)
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
					catch (COMException)
					{
					}

					byte[] metaDataBytes = encoding.GetBytes(metaData);

					if (attachments == null)
					{
						attachments = metaDataBytes;
					}
					else
					{
						attachments = BitBytes.MergeByteArrays(
							attachments, metaDataBytes);
					}

					string filePath = basePath + attachment.FileName;
					attachment.SaveAsFile(filePath);

					byte[] fileBytes = File.ReadAllBytes(filePath);

					attachments =
						BitBytes.MergeByteArrays(attachments, fileBytes);

					Marshal.ReleaseComObject(attachment);
				}
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Warn(exception.ToString());
			}

			return attachments;
		}

		private static byte[] GetBody(MailItem mailItem)
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

				allBody = BitBytes.MergeByteArrays(bodyBytes, htmlBodyBytes);
				allBody = BitBytes.MergeByteArrays(allBody, rtfBody);
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Warn(exception.ToString());
			}

			return allBody;
		}

		private static ushort GetBooleans(MailItem mailItem)
		{
			ushort boolHolder = 0;

			try
			{
				bool rawValue = false;

				try
				{
					rawValue = mailItem.AlternateRecipientAllowed;
				}
				catch (COMException)
				{
				}

				boolHolder = BitBytes.SetBit(boolHolder, 0, rawValue);

				rawValue = mailItem.AutoForwarded;
				boolHolder = BitBytes.SetBit(boolHolder, 1, rawValue);

				rawValue = mailItem.AutoResolvedWinner;
				boolHolder = BitBytes.SetBit(boolHolder, 2, rawValue);

				rawValue = mailItem.DeleteAfterSubmit;
				boolHolder = BitBytes.SetBit(boolHolder, 3, rawValue);

				rawValue = mailItem.IsMarkedAsTask;
				boolHolder = BitBytes.SetBit(boolHolder, 4, rawValue);

				rawValue = mailItem.NoAging;
				boolHolder = BitBytes.SetBit(boolHolder, 5, rawValue);

				rawValue = mailItem.OriginatorDeliveryReportRequested;
				boolHolder = BitBytes.SetBit(boolHolder, 6, rawValue);

				rawValue = mailItem.ReadReceiptRequested;
				boolHolder = BitBytes.SetBit(boolHolder, 7, rawValue);

				rawValue = mailItem.RecipientReassignmentProhibited;
				boolHolder = BitBytes.SetBit(boolHolder, 8, rawValue);

				rawValue = mailItem.ReminderOverrideDefault;
				boolHolder = BitBytes.SetBit(boolHolder, 9, rawValue);

				rawValue = mailItem.ReminderPlaySound;
				boolHolder = BitBytes.SetBit(boolHolder, 10, rawValue);

				rawValue = mailItem.ReminderSet;
				boolHolder = BitBytes.SetBit(boolHolder, 11, rawValue);

				rawValue = mailItem.Saved;
				boolHolder = BitBytes.SetBit(boolHolder, 12, rawValue);

				rawValue = mailItem.Sent;
				boolHolder = BitBytes.SetBit(boolHolder, 13, rawValue);

				rawValue = mailItem.Submitted;
				boolHolder = BitBytes.SetBit(boolHolder, 14, rawValue);

				rawValue = mailItem.UnRead;
				boolHolder = BitBytes.SetBit(boolHolder, 15, rawValue);
			}
			catch (COMException exception)
			{
				Log.Warn(exception.ToString());
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

		private static byte[] GetDateTimes(MailItem mailItem)
		{
			byte[] data = null;

			try
			{
				DateTime deferredDeliveryTimeDateTime = DateTime.MinValue;

				try
				{
					deferredDeliveryTimeDateTime =
						mailItem.DeferredDeliveryTime;
				}
				catch (COMException)
				{
				}

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
				exception is COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Warn(exception.ToString());
			}

			return data;
		}

		private static byte[] GetEnums(string path, MailItem mailItem)
		{
			byte[] buffer = null;

			try
			{
				int bodyFormat = 0;

				try
				{
					bodyFormat = (int)mailItem.BodyFormat;
				}
				catch (COMException)
				{
				}

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
				catch (COMException)
				{
				}

				// 9 ints * size of int
				int bufferSize = 9 * 4;
				buffer = new byte[bufferSize];

				int index = 0;
				buffer =
					BitBytes.CopyIntToByteArray(buffer, index, bodyFormat);
				index += 4;
				buffer = BitBytes.CopyIntToByteArray(buffer, index, itemClass);
				index += 4;
				buffer =
					BitBytes.CopyIntToByteArray(buffer, index, importance);
				index += 4;
				buffer = BitBytes.CopyIntToByteArray(
					buffer, index, markForDownload);
				index += 4;
				buffer =
					BitBytes.CopyIntToByteArray(buffer, index, permission);
				index += 4;
				buffer = BitBytes.CopyIntToByteArray(
					buffer, index, permissionService);
				index += 4;
				buffer =
					BitBytes.CopyIntToByteArray(buffer, index, sensitivity);
				index += 4;
				buffer = BitBytes.CopyIntToByteArray(
					buffer, index, mailItem.InternetCodepage);
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				string sentOn = mailItem.SentOn.ToString(
					"yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

				Log.Error("Exception at: " + path);

				LogFormatMessage.Error(
					"Item: {0}: From: {1}: {2} Subject: {3}",
					sentOn,
					mailItem.SenderName,
					mailItem.SenderEmailAddress,
					mailItem.Subject);

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
					ushort booleans = GetBooleans(mailItem);

					byte[] actions = GetActions(mailItem);
					byte[] attachments = GetAttachments(mailItem);
					byte[] dateTimes = GetDateTimes(mailItem);
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

						Log.Warn("Exception on RTFBody at: " + path);

						LogFormatMessage.Warn(
							"Item: {0}: From: {1}: {2} Subject: {3}",
							sentOn,
							mailItem.SenderName,
							mailItem.SenderEmailAddress,
							mailItem.Subject);
					}

					if (rtfBody != null)
					{
						rtfBody = RtfEmail.Trim(rtfBody);
					}

					byte[] strings = GetStringProperties(mailItem);
					byte[] userProperties = GetUserProperties(mailItem);

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
					long currentIndex = BitBytes.ArrayCopyConditional(
						actions, ref finalBuffer, 0);

					currentIndex = BitBytes.ArrayCopyConditional(
						attachments, ref finalBuffer, currentIndex);
					currentIndex = BitBytes.ArrayCopyConditional(
						dateTimes, ref finalBuffer, currentIndex);
					currentIndex = BitBytes.ArrayCopyConditional(
						enums, ref finalBuffer, currentIndex);
					currentIndex = BitBytes.ArrayCopyConditional(
						rtfBody, ref finalBuffer, currentIndex);
					currentIndex = BitBytes.ArrayCopyConditional(
						strings, ref finalBuffer, currentIndex);
					currentIndex = BitBytes.ArrayCopyConditional(
						userProperties, ref finalBuffer, currentIndex);

					finalBuffer = BitBytes.CopyUshortToByteArray(
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
		private static string GetRecipients(MailItem mailItem)
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

		private static byte[] GetStringProperties(MailItem mailItem)
		{
			byte[] data = null;

			try
			{
				string bcc = mailItem.BCC;
				string billingInformation = null;

				try
				{
					billingInformation = mailItem.BillingInformation;
				}
				catch (COMException)
				{
				}

				string body = mailItem.Body;

				if (body != null)
				{
					body = body.TrimEnd();
				}

				string categories = mailItem.Categories;
				string cc = mailItem.CC;
				string companies = mailItem.Companies;
				string conversationID = mailItem.ConversationID;
				string conversationTopic = mailItem.ConversationTopic;
				string flagRequest = mailItem.FlagRequest;
				string header = mailItem.PropertyAccessor.GetProperty(
					"http://schemas.microsoft.com/mapi/proptag/0x007D001F");

				if (header != null)
				{
					header = RemoveMimeOleVersion(header);
				}

				string htmlBody = mailItem.HTMLBody;

				if (htmlBody != null)
				{
					htmlBody = HtmlEmail.Trim(htmlBody);
				}

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
				exception is COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Error(exception.ToString());
			}

			return data;
		}

		private static byte[] GetUserProperties(MailItem mailItem)
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
						properties = BitBytes.MergeByteArrays(
							properties, metaDataBytes);
					}

					Marshal.ReleaseComObject(property);
				}
			}
			catch (System.Exception exception) when
				(exception is ArgumentException ||
				exception is ArgumentNullException ||
				exception is ArgumentOutOfRangeException ||
				exception is ArrayTypeMismatchException ||
				exception is COMException ||
				exception is InvalidCastException ||
				exception is RankException)
			{
				Log.Warn(exception.ToString());
			}

			return properties;
		}

		private static void LogException(
			string path, string extraInformation, MailItem mailItem)
		{
			string sentOn = mailItem.SentOn.ToString(
				"yyyy-MM-dd HH:mm:ss",
				CultureInfo.InvariantCulture);

			Log.Error("Exception " + extraInformation + "at: " + path);

			LogFormatMessage.Error(
				"Item: {0}: From: {1}: {2} Subject: {3}",
				sentOn,
				mailItem.SenderName,
				mailItem.SenderEmailAddress,
				mailItem.Subject);
		}
	}
}
