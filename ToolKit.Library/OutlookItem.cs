/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookItem.cs" company="James John McGuire">
// Copyright © 2021 - 2024 James John McGuire. All Rights Reserved.
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
	/// Content Item.
	/// </summary>
	public class OutlookItem : IContentItem
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private readonly object mapiItem;
		private readonly string synopses;

		private string hash;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookItem"/> class.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		public OutlookItem(object mapiItem)
		{
			this.mapiItem = mapiItem;

			synopses = GetItemSynopses();
		}

		/// <summary>
		/// Gets the item's hash text.
		/// </summary>
		/// <value>The item's hash stext.</value>
		public string Hash
		{
			get
			{
				if (hash == null)
				{
					hash = GetItemHash();
				}

				return hash;
			}
		}

		/// <summary>
		/// Gets the item's synopses text.
		/// </summary>
		/// <value>The item's synopses text.</value>
		public string Synopses
		{
			get { return synopses; }
		}

		/// <summary>
		/// Deletes the given item.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		public static void Delete(object mapiItem)
		{
			try
			{
				switch (mapiItem)
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
						if (mapiItem != null)
						{
							message += ": " + mapiItem.ToString();
						}

						Log.Warn(message);
						break;
				}

				Marshal.ReleaseComObject(mapiItem);
			}
			catch (COMException exception)
			{
				Log.Error(exception.ToString());
			}
		}

		/// <summary>
		/// Get Actions Data.
		/// </summary>
		/// <param name="actions">The item actions.</param>
		/// <returns>The item actions data.</returns>
		public static byte[] GetActions(Actions actions)
		{
			byte[] actionsData = null;

			if (actions != null)
			{
				int total = actions.Count;

				for (int index = 1; index <= total; index++)
				{
					Microsoft.Office.Interop.Outlook.Action action =
						actions[index];

					byte[] metaDataBytes = GetActionData(action);

					if (actionsData == null)
					{
						actionsData = metaDataBytes;
					}
					else
					{
						actionsData =
							BitBytes.MergeByteArrays(actionsData, metaDataBytes);
					}

					Marshal.ReleaseComObject(action);
				}
			}

			return actionsData;
		}

		/// <summary>
		/// Get Attachments Data.
		/// </summary>
		/// <param name="attachments">The item attachments.</param>
		/// <returns>The item attachments data.</returns>
		public static byte[] GetAttachments(Attachments attachments)
		{
			byte[] attachmentsData = null;

			if (attachments != null)
			{
				int total = attachments.Count;

				for (int index = 1; index <= total; index++)
				{
					Attachment attachment = attachments[index];

					byte[] attachementData = GetAttachmentData(attachment);

					if (attachmentsData == null)
					{
						attachmentsData = attachementData;
					}
					else
					{
						attachmentsData = BitBytes.MergeByteArrays(
							attachmentsData, attachementData);
					}

					Marshal.ReleaseComObject(attachment);
				}
			}

			return attachmentsData;
		}

		/// <summary>
		/// Get DateTime Properites Data.
		/// </summary>
		/// <param name="times">The DataTime properties data.</param>
		/// <returns>The DataTime properties data in bytes.</returns>
		public static byte[] GetDateTimesBytes(IList<DateTime> times)
		{
			byte[] data = null;

			if (times != null)
			{
				List<string> timesStrings = [];

				foreach (DateTime time in times)
				{
					string timeString = time.ToString("O");
					timesStrings.Add(timeString);
				}

				StringBuilder builder = new ();

				foreach (string timeString in timesStrings)
				{
					builder.Append(timeString);
				}

				string buffer = builder.ToString();

				Encoding encoding = Encoding.UTF8;
				data = encoding.GetBytes(buffer);
			}

			return data;
		}

		/// <summary>
		/// Get enums properites data.
		/// </summary>
		/// <param name="ints">The enums properties data.</param>
		/// <returns>The enums properties data in bytes.</returns>
		public static byte[] GetEnumsBuffer(IList<int> ints)
		{
			byte[] buffer = null;

			if (ints != null)
			{
				int bufferSize = ints.Count * 4;
				buffer = new byte[bufferSize];

				int index = 0;
				foreach (int item in ints)
				{
					buffer = BitBytes.CopyIntToByteArray(buffer, index, item);
					index += 4;
				}
			}

			return buffer;
		}

		/// <summary>
		/// Get recipients properites data.
		/// </summary>
		/// <param name="recipients">The enums recipients data.</param>
		/// <returns>The recipients properties data in bytes.</returns>
		[System.Diagnostics.CodeAnalysis.SuppressMessage(
			"StyleCop.CSharp.NamingRules",
			"SA1305:Field names should not use Hungarian notation",
			Justification = "It isn't hungarian notation.")]
		public static byte[] GetRecipients(Recipients recipients)
		{
			byte[] data = null;

			if (recipients != null)
			{
				string recipientsData;
				List<string> toList = [];
				List<string> ccList = [];
				List<string> bccList = [];

				int total = recipients.Count;

				for (int index = 1; index <= total; index++)
				{
					Recipient recipient = recipients[index];
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

				StringBuilder builder = new ();

				foreach (string formattedRecipient in toList)
				{
					builder.Append(formattedRecipient);
				}

				foreach (string formattedRecipient in ccList)
				{
					builder.Append(formattedRecipient);
				}

				foreach (string formattedRecipient in bccList)
				{
					builder.Append(formattedRecipient);
				}

				recipientsData = builder.ToString();

				Encoding encoding = Encoding.UTF8;
				data = encoding.GetBytes(recipientsData);
			}

			return data;
		}

		/// <summary>
		/// Get user properites data.
		/// </summary>
		/// <param name="userProperties">The user properties data.</param>
		/// <returns>The user properties data in bytes.</returns>
		public static byte[] GetUserProperties(UserProperties userProperties)
		{
			byte[] properties = null;

			if (userProperties != null)
			{
				int total = userProperties.Count;

				for (int index = 1; index <= total; index++)
				{
					UserProperty property = userProperties[index];
					properties = GetUserProperty(properties, property);
				}
			}

			return properties;
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		/// <param name="destination">The destination folder.</param>
		public static void Move(object mapiItem, MAPIFolder destination)
		{
			try
			{
				switch (mapiItem)
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
						mailItem = mailItem.Move(destination);
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
						if (mapiItem != null)
						{
							message += ": " + mapiItem.ToString();
						}

						Log.Warn(message);
						break;
				}

				Marshal.ReleaseComObject(mapiItem);
			}
			catch (COMException exception)
			{
				Log.Error(exception.ToString());
			}
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		/// <param name="destination">The destination folder.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public static async Task MoveAsync(
			object mapiItem, MAPIFolder destination)
		{
			CancellationTokenSource source = new ();

			try
			{
				source.CancelAfter(TimeSpan.FromSeconds(5));

				switch (mapiItem)
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
							mailItem = mailItem.Move(destination)).
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
						if (mapiItem != null)
						{
							message += ": " + mapiItem.ToString();
						}

						Log.Warn(message);
						break;
				}

				Marshal.ReleaseComObject(mapiItem);
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

		/// <summary>
		/// Deletes the given item.
		/// </summary>
		public void Delete()
		{
			Delete(mapiItem);
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="destination">The destination folder.</param>
		public void Move(MAPIFolder destination)
		{
			Move(mapiItem, destination);
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="destination">The destination folder.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MoveAsync(MAPIFolder destination)
		{
			await MoveAsync(mapiItem, destination).ConfigureAwait(false);
		}

		private static byte[] GetActionData(
			Microsoft.Office.Interop.Outlook.Action action)
		{
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

			return metaDataBytes;
		}

		private static byte[] GetAttachmentData(Attachment attachment)
		{
			string basePath = Path.GetTempPath();

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
				"{0}{1}{2}{3}",
				attachment.DisplayName,
				indexValue,
				position,
				attachmentType);

			try
			{
				metaData += attachment.FileName;
			}
			catch (COMException)
			{
			}

			try
			{
				metaData += attachment.PathName;
			}
			catch (COMException)
			{
			}

			byte[] metaDataBytes = encoding.GetBytes(metaData);

			string filePath = basePath + attachment.FileName;
			attachment.SaveAsFile(filePath);

			byte[] fileBytes = File.ReadAllBytes(filePath);

			byte[] attachmentData =
				BitBytes.MergeByteArrays(metaDataBytes, fileBytes);

			return attachmentData;
		}

		private static long GetBufferSize(IList<byte[]> buffers)
		{
			long bufferSize = 0;

			foreach (byte[] buffer in buffers)
			{
				if (buffer != null)
				{
					bufferSize += buffer.LongLength;
				}
			}

			bufferSize += 2;

			return bufferSize;
		}

		private static byte[] GetUserProperty(
			byte[] properties, UserProperty property)
		{
			byte[] userPropertyData = GetUserPropertyData(property);

			if (properties == null)
			{
				properties = userPropertyData;
			}
			else
			{
				properties = BitBytes.MergeByteArrays(
					properties, userPropertyData);
			}

			Marshal.ReleaseComObject(property);

			return properties;
		}

		private static byte[] GetUserPropertyData(UserProperty property)
		{
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

			return metaDataBytes;
		}

		private string GetPath()
		{
			string path = null;

			if (mapiItem != null)
			{
				try
				{
					MAPIFolder parent = null;

					switch (mapiItem)
					{
						case AppointmentItem appointmentItem:
							parent = appointmentItem.Parent;
							break;
						case MailItem mailItem:
							parent = mailItem.Parent;
							break;
						default:
							string message = "Item is of unsupported type: " +
								mapiItem.ToString();
							Log.Warn(message);
							break;
					}

					path = OutlookFolder.GetFolderPath(parent);
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
			}

			return path;
		}

		private void LogException()
		{
			string path = GetPath();
			Log.Error("Exception at: " + path);

			LogFormatMessage.Error("Item: {0}:", synopses);
		}

		private byte[] GetItemBytes(bool strict = false)
		{
			byte[] itemBytes = null;

			if (mapiItem != null)
			{
				try
				{
					IList<byte[]> buffers = [];

					switch (mapiItem)
					{
						case AppointmentItem appointmentItem:
							OutlookAppointment appointment = new (mapiItem);
							buffers = appointment.GetProperties(strict);
							break;
						case MailItem mailItem:
							OutlookMail mail = new (mapiItem);
							buffers = mail.GetProperties(strict);
							break;
						default:
							string message = "Item is of unsupported type: " +
								mapiItem.ToString();
							Log.Warn(message);
							break;
					}

					long bufferSize = GetBufferSize(buffers);

					itemBytes = new byte[bufferSize];

					// combine the parts
					long currentIndex = 0;
					foreach (byte[] buffer in buffers)
					{
						currentIndex = BitBytes.ArrayCopyConditional(
							buffer, ref itemBytes, currentIndex);
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
					Log.Error(exception.ToString());
				}
			}

			return itemBytes;
		}

		private string GetItemHash()
		{
			string hashBase64 = null;

			if (mapiItem != null)
			{
				try
				{
					byte[] itemBytes = GetItemBytes();

					if (itemBytes != null)
					{
#if NET5_0_OR_GREATER
						byte[] hashValue = SHA256.HashData(itemBytes);
#else
					using SHA256 hasher = SHA256.Create();
					byte[] hashValue = hasher.ComputeHash(itemBytes);
#endif
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
					LogException();
					Log.Error(exception.ToString());
				}
			}

			return hashBase64;
		}

		private string GetItemSynopses()
		{
			string synopses = null;

			if (mapiItem != null)
			{
				try
				{
					switch (mapiItem)
					{
						case AppointmentItem appointmentItem:
							synopses = OutlookAppointment.GetSynopses(appointmentItem);
							break;
						case MailItem mailItem:
							synopses = OutlookMail.GetSynopses(mailItem);
							break;
						default:
							string message = "Item is of unsupported type: " +
								mapiItem.ToString();
							Log.Warn(message);
							break;
					}
				}
				catch (COMException exception)
				{
					Log.Error(exception.ToString());
				}
			}

			return synopses;
		}
	}
}
