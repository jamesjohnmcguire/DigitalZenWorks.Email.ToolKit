/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookMail.cs" company="James John McGuire">
// Copyright © 2021 - 2024 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using DigitalZenWorks.Common.Utilities;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Outlook Mail Class.
	/// </summary>
	public class OutlookMail
	{
		private readonly MailItem mailItem;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookMail"/> class.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		public OutlookMail(object mapiItem)
		{
			ArgumentNullException.ThrowIfNull(mapiItem);

			mailItem = mapiItem as MailItem;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <param name="mailItem">The MailItem to check.</param>
		/// <returns>The synoses of the item.</returns>
		public static string GetSynopses(MailItem mailItem)
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
		/// Get the bytes of all relevant properties.
		/// </summary>
		/// <param name="strict">Indicates whether the check should be strict
		/// or not.</param>
		/// <returns>The bytes of all relevant properties.</returns>
		public IList<byte[]> GetProperties(bool strict = false)
		{
			List<byte[]> buffers = [];
			byte[] actions = null;
			byte[] attachments = null;
			ushort booleans = 0;
			byte[] dateTimes = null;
			byte[] enums = null;
			byte[] recipients = null;
			byte[] strings = null;
			byte[] userProperties = null;

			booleans = GetBooleans();

			actions = ContentItem.GetActions(mailItem.Actions);
			buffers.Add(actions);

			attachments = ContentItem.GetAttachments(
				mailItem.Attachments);
			buffers.Add(attachments);

			dateTimes = GetDateTimes();
			buffers.Add(dateTimes);

			enums = GetEnums();
			buffers.Add(enums);

			recipients = ContentItem.GetRecipients(mailItem.Recipients);
			buffers.Add(recipients);

			strings = GetStringProperties(strict);
			buffers.Add(strings);

			userProperties = ContentItem.GetUserProperties(
				mailItem.UserProperties);
			buffers.Add(userProperties);

			byte[] itemBytes = new byte[2];
			itemBytes = BitBytes.CopyUshortToByteArray(
				itemBytes, 0, booleans);
			buffers.Add(itemBytes);

			return buffers;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <returns>The synoses of the item.</returns>
		public string GetSynopses()
		{
			string synopses = GetSynopses(mailItem);

			return synopses;
		}

		private static string NormalizeHeaders(string headers)
		{
#if NETCOREAPP1_0_OR_GREATER
			string[] parts = headers.Split("\r\n");
#else
			string[] parts = headers.Split('\n');
#endif

			List<string> list = new (parts);

			list.Sort();

#if NETCOREAPP1_0_OR_GREATER
			headers = string.Join("\r\n", list);
#else
			headers = string.Join("\n", list);
#endif

			return headers;
		}

		private ushort GetBooleans()
		{
			ushort boolHolder = 0;

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

			return boolHolder;
		}

		private byte[] GetDateTimes()
		{
			byte[] data = null;

			List<DateTime> times = [];

			DateTime deferredDeliveryTimeDateTime = DateTime.MinValue;

			try
			{
				deferredDeliveryTimeDateTime =
					mailItem.DeferredDeliveryTime;
			}
			catch (COMException)
			{
			}

			times.Add(deferredDeliveryTimeDateTime);

			DateTime expiryTimeDateTime = mailItem.ExpiryTime;
			times.Add(expiryTimeDateTime);

			DateTime receivedTimeDateTime = mailItem.ReceivedTime;
			times.Add(receivedTimeDateTime);

			DateTime reminderTimeDateTime = mailItem.ReminderTime;
			times.Add(reminderTimeDateTime);

			DateTime retentionExpirationDateDateTime =
				mailItem.RetentionExpirationDate;
			times.Add(retentionExpirationDateDateTime);

			DateTime sentOnDateTime = mailItem.SentOn;
			times.Add(sentOnDateTime);

			DateTime taskCompletedDateDateTime =
				mailItem.TaskCompletedDate;
			times.Add(taskCompletedDateDateTime);

			DateTime taskDueDateDateTime = mailItem.TaskDueDate;
			times.Add(taskDueDateDateTime);

			DateTime taskStartDateDateTime = mailItem.TaskStartDate;
			times.Add(taskStartDateDateTime);

			data = ContentItem.GetDateTimesBytes(times);

			return data;
		}

		private byte[] GetEnums()
		{
			byte[] buffer = null;

			List<int> ints = [];

			int bodyFormat = 0;

			try
			{
				bodyFormat = (int)mailItem.BodyFormat;
			}
			catch (COMException)
			{
			}

			ints.Add(bodyFormat);

			int itemClass = (int)mailItem.Class;
			ints.Add(itemClass);

			int importance = (int)mailItem.Importance;
			ints.Add(importance);

			int markForDownload = (int)mailItem.MarkForDownload;
			ints.Add(markForDownload);

			int permission = 0;

			try
			{
				permission = (int)mailItem.Permission;
			}
			catch (COMException)
			{
			}

			ints.Add(permission);

			int permissionService = (int)mailItem.PermissionService;
			ints.Add(permissionService);

			int sensitivity = (int)mailItem.Sensitivity;
			ints.Add(sensitivity);

			buffer = ContentItem.GetEnumsBuffer(ints);

			return buffer;
		}

		private byte[] GetStringProperties(
			bool strict = false,
			bool ignoreConversation = true)
		{
			byte[] data = null;

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

			if (body != null && strict == false)
			{
				body = body.TrimEnd();
			}

			string categories = mailItem.Categories;
			string cc = mailItem.CC;
			string companies = mailItem.Companies;
			string conversationID = null;

			if (ignoreConversation == false)
			{
				conversationID = mailItem.ConversationID;
			}

			string conversationTopic = mailItem.ConversationTopic;
			string flagRequest = mailItem.FlagRequest;
			string header = mailItem.PropertyAccessor.GetProperty(
				"http://schemas.microsoft.com/mapi/proptag/0x007D001F");

			if (header != null && strict == false)
			{
				header = ContentItem.RemoveMimeOleVersion(header);

#if NETCOREAPP1_0_OR_GREATER
				header = header.Replace(
					"Errors-to:",
					"Errors-To:",
					StringComparison.Ordinal);
#else
				header = header.Replace(
					"Errors-to:",
					"Errors-To:");
#endif

				header = NormalizeHeaders(header);
			}

			string htmlBody = mailItem.HTMLBody;

			if (htmlBody != null && strict == false)
			{
				htmlBody = HtmlEmail.Trim(htmlBody);
			}

			string messageClass = mailItem.MessageClass;
			string mileage = mailItem.Mileage;
			string receivedByEntryID = null;
			string receivedByName = mailItem.ReceivedByName;
			string receivedOnBehalfOfEntryID = null;

			string receivedOnBehalfOfName = null;
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

			if (strict == true)
			{
				// Might need to investigate further.
				receivedByEntryID = mailItem.ReceivedByEntryID;
				receivedOnBehalfOfEntryID =
					mailItem.ReceivedOnBehalfOfEntryID;
				receivedOnBehalfOfName = mailItem.ReceivedOnBehalfOfName;
			}

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

			return data;
		}
	}
}
