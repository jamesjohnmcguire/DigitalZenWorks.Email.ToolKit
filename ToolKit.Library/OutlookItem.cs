/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookItem.cs" company="James John McGuire">
// Copyright © 2021 - 2025 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Common.Utilities;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

			synopses = GetSynopses();
		}

		/// <summary>
		/// Gets the item's hash text.
		/// </summary>
		/// <value>The item's hash stext.</value>
		public string Hash
		{
			get
			{
				hash ??= GetItemHash();

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
		/// Delete duplicate item.
		/// </summary>
		/// <param name="session">The Outlook session.</param>
		/// <param name="duplicateId">The duplicate id.</param>
		/// <param name="keeperSynopses">The keeper synopses.</param>
		/// <param name="dryRun">Indicates if this is a dry run or not.</param>
		public static void DeleteDuplicate(
			NameSpace session,
			string duplicateId,
			string keeperSynopses,
			bool dryRun)
		{
			if (session != null)
			{
				try
				{
					object mapiItem = session.GetItemFromID(duplicateId);

					if (mapiItem != null)
					{
						bool isValidDuplicate =
							DoubleCheckDuplicate(keeperSynopses, mapiItem);

						if (isValidDuplicate == true && dryRun == false)
						{
							OutlookItem contentItem = new (mapiItem);
							contentItem.Delete();
						}

						Marshal.ReleaseComObject(mapiItem);
					}
				}
				catch (System.Exception exception) when
				(exception is COMException ||
				exception is InvalidCastException)
				{
					Log.Error(exception.ToString());
				}
			}
		}

		/// <summary>
		/// Format for text an enum value.
		/// </summary>
		/// <param name="propertyName">The enum property name.</param>
		/// <param name="propertyValue">The enum property value.</param>
		/// <returns>A formatted string of the enum value.</returns>
		public static string FormatEnumValue(
			string propertyName,
			object propertyValue)
		{
			int enumValue = (int)propertyValue;
			string textValue =
				enumValue.ToString(CultureInfo.InvariantCulture);

			string result = string.Format(
					CultureInfo.InvariantCulture,
					"{0}: {1}",
					propertyName,
					textValue);

			result += Environment.NewLine;

			return result;
		}

		/// <summary>
		/// Format property value.
		/// </summary>
		/// <param name="propertyName">The property name.</param>
		/// <param name="propertyValue">The property value.</param>
		/// <returns>A formatted string of the property value.</returns>
		public static string FormatValue(
			string propertyName,
			string propertyValue)
		{
			string formattedText = null;

			if (propertyValue != null)
			{
				formattedText = string.Format(
						CultureInfo.InvariantCulture,
						"{0}: {1}",
						propertyName,
						propertyValue);

				formattedText += Environment.NewLine;
			}

			return formattedText;
		}

		/// <summary>
		/// Format value on condition.
		/// </summary>
		/// <param name="propertyName">The property name.</param>
		/// <param name="propertyValue">The property value.</param>
		/// <param name="formatValue">The value on whether to format
		/// or not.</param>
		/// <returns>A formatted string of the property value.</returns>
		public static string FormatValueConditional(
			string propertyName,
			string propertyValue,
			bool formatValue = false)
		{
			if (formatValue == true)
			{
				propertyValue = FormatValue(propertyName, propertyValue);
			}

			return propertyValue;
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
				string actionsText = GetActionsText(actions);

				Encoding encoding = Encoding.UTF8;
				actionsData = encoding.GetBytes(actionsText);
			}

			return actionsData;
		}

		/// <summary>
		/// Get Actions Data.
		/// </summary>
		/// <param name="actions">The item actions.</param>
		/// <param name="addNewLine">Indicates whether to add a new line
		/// or not.</param>
		/// <returns>The item actions data as text.</returns>
		public static string GetActionsText(
			Actions actions, bool addNewLine = false)
		{
			string actionsText = null;

			if (actions != null)
			{
				int total = actions.Count;

				for (int index = 1; index <= total; index++)
				{
					Microsoft.Office.Interop.Outlook.Action action =
						actions[index];

					string actionText = GetActionText(action, addNewLine);

					if (actionText == null)
					{
						actionsText = actionText;
					}
					else
					{
						actionsText += actionText;
					}

					if (addNewLine == true)
					{
						actionsText += Environment.NewLine;
					}

					Marshal.ReleaseComObject(action);
				}
			}

			return actionsText;
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
				string attachmentsText = GetAttachmentsText(attachments);

				if (attachmentsText != null)
				{
					Encoding encoding = Encoding.UTF8;
					attachmentsData = encoding.GetBytes(attachmentsText);
				}
			}

			return attachmentsData;
		}

		/// <summary>
		/// Get Attachments Data.
		/// </summary>
		/// <param name="attachments">The item attachments.</param>
		/// <returns>The item attachments data as text.</returns>
		public static string GetAttachmentsText(Attachments attachments)
		{
			string attachmentsText = null;

			if (attachments != null)
			{
				int total = attachments.Count;

				for (int index = 1; index <= total; index++)
				{
					Attachment attachment = attachments[index];

					string attachmentText = GetAttachmentText(attachment);

					if (attachmentsText == null)
					{
						attachmentsText = attachmentText;
					}
					else
					{
						attachmentsText += attachmentText;
					}

					Marshal.ReleaseComObject(attachment);
				}
			}

			return attachmentsText;
		}

		/// <summary>
		/// Get boolean text.
		/// </summary>
		/// <param name="item">The item to check.</param>
		/// <returns>The formatted text value.</returns>
		public static string GetBooleanText(bool item)
		{
			string name = nameof(item);

			string booleanText = string.Format(
				CultureInfo.InvariantCulture,
				"{0}: {1}",
				name,
				item.ToString());

			return booleanText;
		}

		/// <summary>
		/// Get DateTime Properites Data.
		/// </summary>
		/// <param name="times">The DataTime properties data.</param>
		/// <returns>The DataTime properties data in bytes.</returns>
		[Obsolete("GetDateTimesBytes(List<DateTime> is Deprecated. " +
			"Use the overload accepting ReadOnlyCollection<DateTime> instead.")]
		[System.Diagnostics.CodeAnalysis.SuppressMessage(
			"Microsoft.Design",
			"CA1002:AvoidExcessiveList",
			Justification = "It Is a Current API.")]
		public static byte[] GetDateTimesBytes(List<DateTime> times)
		{
			byte[] data = null;

			if (times != null)
			{
				ReadOnlyCollection<DateTime> timesReadOnly = new (times);
				string buffer = GetDateTimesText(timesReadOnly);

				Encoding encoding = Encoding.UTF8;
				data = encoding.GetBytes(buffer);
			}

			return data;
		}

		/// <summary>
		/// Get DateTime Properites Data.
		/// </summary>
		/// <param name="times">The DataTime properties data.</param>
		/// <returns>The DataTime properties data in bytes.</returns>
		public static byte[] GetDateTimesBytes(
			ReadOnlyCollection<DateTime> times)
		{
			byte[] data = null;

			if (times != null)
			{
				string buffer = GetDateTimesText(times);

				Encoding encoding = Encoding.UTF8;
				data = encoding.GetBytes(buffer);
			}

			return data;
		}

		/// <summary>
		/// Get DateTime Properites Data.
		/// </summary>
		/// <param name="times">The DataTime properties data.</param>
		/// <param name="labels">A list of lables.</param>
		/// <returns>The DataTime properties data as text.</returns>
		[Obsolete("GetDateTimesText(List<DateTime>, " +
			"ReadOnlyCollection<string> is Deprecated." +
			"Use the overload accepting ReadOnlyCollection<DateTime> instead.")]
		[System.Diagnostics.CodeAnalysis.SuppressMessage(
			"Microsoft.Design",
			"CA1002:AvoidExcessiveList",
			Justification = "It Is a Current API.")]
		public static string GetDateTimesText(
			List<DateTime> times, ReadOnlyCollection<string> labels = null)
		{
			ReadOnlyCollection<DateTime> timesReadOnly = new (times);
			string dateTimesText = GetDateTimesText(timesReadOnly, labels);

			return dateTimesText;
		}

		/// <summary>
		/// Get DateTime Properites Data.
		/// </summary>
		/// <param name="times">The DataTime properties data.</param>
		/// <param name="labels">A list of lables.</param>
		/// <returns>The DataTime properties data as text.</returns>
		public static string GetDateTimesText(
			ReadOnlyCollection<DateTime> times,
			ReadOnlyCollection<string> labels = null)
		{
			string dateTimesText = null;

			if (times != null)
			{
				StringBuilder builder = new ();

				for (int index = 0; index < times.Count; index++)
				{
					DateTime time = times[index];

					string timeString = time.ToString("O");

					if (labels != null)
					{
						string label = labels[index];
						timeString = FormatValue(label, timeString);
					}

					builder.Append(timeString);
				}

				dateTimesText = builder.ToString();
			}

			return dateTimesText;
		}

		/// <summary>
		/// Get details.
		/// </summary>
		/// <param name="mapiItem">The item to inspect.</param>
		/// <param name="strict">Indicates whether to use a strict check
		/// or not.</param>
		/// <returns>A text of the details.</returns>
		public static string GetDetails(
			object mapiItem, bool strict = false)
		{
			string details = null;

			if (mapiItem != null)
			{
				try
				{
					IList<byte[]> buffers = [];

					switch (mapiItem)
					{
						case AppointmentItem appointmentItem:
							OutlookAppointment appointment = new (mapiItem);
							details = appointment.GetPropertiesText(strict);
							break;
						case ContactItem contact:
							OutlookContact outlookContact = new (mapiItem);
							details = outlookContact.GetPropertiesText(strict);
							break;
						case MailItem mailItem:
							OutlookMail mail = new (mapiItem);
							details = mail.GetPropertiesText(strict, true);
							break;
						default:
							string message = "Item is of unsupported type: " +
								mapiItem.ToString();
							Log.Warn(message);
							break;
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

			return details;
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
		/// Gets the item's hash.
		/// </summary>
		/// <param name="mapiItem">The items to compute.</param>
		/// <returns>The item's hash encoded in base 64.</returns>
		public static async Task<string> GetHashAsync(object mapiItem)
		{
			string hashBase64 = null;

			if (mapiItem != null)
			{
				try
				{
					byte[] itemBytes = await Task.Run(() =>
							GetItemBytes(mapiItem)).ConfigureAwait(false);

					if (itemBytes != null)
					{
						hashBase64 = GetBytesHash(itemBytes);
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
					LogException(mapiItem);
					Log.Error(exception.ToString());
				}
			}

			return hashBase64;
		}

		/// <summary>
		/// Get the current item's folder path.
		/// </summary>
		/// <param name="mapiItem">The item to check.</param>
		/// <returns>The current item's folder path.</returns>
		public static string GetPath(object mapiItem)
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

		/// <summary>
		/// Get recipients properites data.
		/// </summary>
		/// <param name="recipients">The enums recipients data.</param>
		/// <returns>The recipients properties data in bytes.</returns>
		public static byte[] GetRecipients(Recipients recipients)
		{
			byte[] data = null;

			if (recipients != null)
			{
				string recipientsData = GetRecipientsText(recipients);

				Encoding encoding = Encoding.UTF8;
				data = encoding.GetBytes(recipientsData);
			}

			return data;
		}

		/// <summary>
		/// Get recipients properites data.
		/// </summary>
		/// <param name="recipients">The enums recipients data.</param>
		/// <returns>The recipients properties data as text.</returns>
		[System.Diagnostics.CodeAnalysis.SuppressMessage(
			"StyleCop.CSharp.NamingRules",
			"SA1305:Field names should not use Hungarian notation",
			Justification = "It isn't hungarian notation.")]
		public static string GetRecipientsText(Recipients recipients)
		{
			string recipientsData = null;

			if (recipients != null)
			{
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
			}

			return recipientsData;
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
				Encoding encoding = Encoding.UTF8;

				string propertiesText = GetUserPropertiesText(userProperties);

				properties = encoding.GetBytes(propertiesText);
			}

			return properties;
		}

		/// <summary>
		/// Get user properites data.
		/// </summary>
		/// <param name="userProperties">The user properties data.</param>
		/// <returns>The user properties data as text.</returns>
		public static string GetUserPropertiesText(
			UserProperties userProperties)
		{
			string properties = null;

			if (userProperties != null)
			{
				properties = string.Empty;
				int total = userProperties.Count;

				for (int index = 1; index <= total; index++)
				{
					UserProperty property = userProperties[index];
					properties += GetUserPropertyText(property);
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

		private static bool DoubleCheckDuplicate(
			string baseSynopses, object mapiItem)
		{
			bool valid = true;
			string duplicateSynopses = GetSynopses(mapiItem);

			if (!duplicateSynopses.Equals(
				baseSynopses, StringComparison.Ordinal))
			{
				Log.Error("Warning! Duplicate Items Don't Seem to Match");
				Log.Error("Not Matching Item: " + duplicateSynopses);

				valid = false;
			}

			return valid;
		}

		private static byte[] GetActionData(
			Microsoft.Office.Interop.Outlook.Action action)
		{
			byte[] actionData = null;

			if (action != null)
			{
				Encoding encoding = Encoding.UTF8;

				string metaData = GetActionText(action);

				actionData = encoding.GetBytes(metaData);
			}

			return actionData;
		}

		private static string GetActionText(
			Microsoft.Office.Interop.Outlook.Action action,
			bool addNewLine = false)
		{
			string actionText = null;

			if (action != null)
			{
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

				string format = "{0}{1}{2}{3}{4}{5}{6}";

				if (addNewLine == true)
				{
					format = "{0} {1} {2} {3} {4} {5} {6}";
				}

				actionText = string.Format(
					CultureInfo.InvariantCulture,
					format,
					copyLike,
					enabled,
					action.Name,
					action.Prefix,
					replyStyle,
					responseStyle,
					showOn);
			}

			return actionText;
		}

		private static byte[] GetAttachmentData(Attachment attachment)
		{
			byte[] attachmentData = null;

			if (attachment != null)
			{
				string attachmentText = GetAttachmentText(attachment);

				Encoding encoding = Encoding.UTF8;
				attachmentData = encoding.GetBytes(attachmentText);
			}

			return attachmentData;
		}

		private static string GetAttachmentText(Attachment attachment)
		{
			string attachmentData = null;

			if (attachment != null)
			{
				string basePath = Path.GetTempPath();

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

				string filePath = basePath + attachment.FileName;
				attachment.SaveAsFile(filePath);

				byte[] fileBytes = File.ReadAllBytes(filePath);
				string hashBase64 = Convert.ToBase64String(fileBytes);

				attachmentData = metaData + hashBase64;
			}

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

		private static string GetBytesHash(byte[] data)
		{
#if NET5_0_OR_GREATER
			byte[] hashValue = SHA256.HashData(data);
#else
			using SHA256 hasher = SHA256.Create();
			byte[] hashValue = hasher.ComputeHash(data);
#endif
			string hashBase64 = Convert.ToBase64String(hashValue);

			return hashBase64;
		}

		private static byte[] GetItemBytes(
			object mapiItem, bool strict = false)
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
						case ContactItem contact:
							OutlookContact outlookContact = new (mapiItem);
							buffers = outlookContact.GetProperties(strict);
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
			byte[] propertyData = null;

			if (property != null)
			{
				string metaData = GetUserPropertyText(property);

				Encoding encoding = Encoding.UTF8;
				propertyData = encoding.GetBytes(metaData);
			}

			return propertyData;
		}

		private static string GetUserPropertyText(UserProperty property)
		{
			string propertyText = null;

			if (property != null)
			{
				int typeEnum = (int)property.Type;
				var propertyValue = property.Value;

				string typeValue =
					typeEnum.ToString(CultureInfo.InvariantCulture);
				string value =
					propertyValue.ToString(CultureInfo.InvariantCulture);

				propertyText = string.Format(
					CultureInfo.InvariantCulture,
					"{0}{1}{2}{3}{4}{5}",
					property.Formula,
					property.Name,
					typeValue,
					property.ValidationFormula,
					property.ValidationText,
					value);
			}

			return propertyText;
		}

		private static string GetSynopses(object mapiItem)
		{
			string synopses = null;

			if (mapiItem != null)
			{
				try
				{
					switch (mapiItem)
					{
						case AppointmentItem appointmentItem:
							synopses = OutlookAppointment.GetSynopses(
								appointmentItem);
							break;
						case ContactItem contactItem:
							synopses = OutlookContact.GetSynopses(
								contactItem);
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

		private static void LogException(object mapiItem)
		{
			string path = GetPath(mapiItem);
			Log.Error("Exception at: " + path);

			string synopses = GetSynopses(mapiItem);
			LogFormatMessage.Error("Item: {0}:", synopses);
		}

		private byte[] GetItemBytes(bool strict = false)
		{
			byte[] itemBytes = GetItemBytes(mapiItem, strict);

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
						hashBase64 = GetBytesHash(itemBytes);
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
					LogException(mapiItem);
					Log.Error(exception.ToString());
				}
			}

			return hashBase64;
		}

		private string GetSynopses()
		{
			string synopses = GetSynopses(mapiItem);

			return synopses;
		}
	}
}
