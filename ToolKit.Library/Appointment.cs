/////////////////////////////////////////////////////////////////////////////
// <copyright file="Appointment.cs" company="James John McGuire">
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
	/// Appointment Class.
	/// </summary>
	public class Appointment
	{
		private readonly AppointmentItem appointmentItem;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="Appointment"/> class.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		public Appointment(object mapiItem)
		{
			appointmentItem = mapiItem as AppointmentItem;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <param name="appointmentItem">The AppointmentItemto check.</param>
		/// <returns>The synoses of the item.</returns>
		public static string GetSynopses(AppointmentItem appointmentItem)
		{
			string synopses = null;

			if (appointmentItem != null)
			{
				string sentOn = appointmentItem.Start.ToString(
					"yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

				synopses = string.Format(
					CultureInfo.InvariantCulture,
					"{0}: From: {1}: {2} Subject: {3}",
					sentOn,
					appointmentItem.Organizer,
					appointmentItem.Subject,
					appointmentItem.Body);
			}

			return synopses;
		}

		/// <summary>
		/// Get the bytes of all relevant properties.
		/// </summary>
		/// <param name="strict">Indicates whether the check should be strict
		/// or not.</param>
		/// <returns>The bytes of all relevant properties.</returns>
		public IList<byte[]> GetPropertiesBytes(bool strict = false)
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

			actions = ContentItem.GetActions(appointmentItem.Actions);
			buffers.Add(actions);

			attachments = ContentItem.GetAttachments(
				appointmentItem.Attachments);
			buffers.Add(attachments);

			dateTimes = GetDateTimes();
			buffers.Add(dateTimes);

			enums = GetEnums();
			buffers.Add(enums);

			recipients = ContentItem.GetRecipients(appointmentItem.Recipients);
			buffers.Add(recipients);

			strings = GetStringProperties(strict);
			buffers.Add(strings);

			userProperties = ContentItem.GetUserProperties(
				appointmentItem.UserProperties);
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
			string synopses = GetSynopses(appointmentItem);

			return synopses;
		}

		private ushort GetBooleans()
		{
			ushort boolHolder = 0;

			bool rawValue = appointmentItem.AllDayEvent;
			boolHolder = BitBytes.SetBit(boolHolder, 0, rawValue);

			rawValue = appointmentItem.AutoResolvedWinner;
			boolHolder = BitBytes.SetBit(boolHolder, 1, rawValue);

			rawValue = appointmentItem.ForceUpdateToAllAttendees;
			boolHolder = BitBytes.SetBit(boolHolder, 2, rawValue);

			rawValue = appointmentItem.IsConflict;
			boolHolder = BitBytes.SetBit(boolHolder, 3, rawValue);

			rawValue = appointmentItem.IsRecurring;
			boolHolder = BitBytes.SetBit(boolHolder, 4, rawValue);

			rawValue = appointmentItem.NoAging;
			boolHolder = BitBytes.SetBit(boolHolder, 5, rawValue);

			rawValue = appointmentItem.ReminderOverrideDefault;
			boolHolder = BitBytes.SetBit(boolHolder, 6, rawValue);

			rawValue = appointmentItem.ReminderPlaySound;
			boolHolder = BitBytes.SetBit(boolHolder, 7, rawValue);

			rawValue = appointmentItem.ReminderSet;
			boolHolder = BitBytes.SetBit(boolHolder, 8, rawValue);

			rawValue = appointmentItem.ResponseRequested;
			boolHolder = BitBytes.SetBit(boolHolder, 9, rawValue);

			rawValue = appointmentItem.Saved;
			boolHolder = BitBytes.SetBit(boolHolder, 10, rawValue);

			rawValue = appointmentItem.UnRead;
			boolHolder = BitBytes.SetBit(boolHolder, 11, rawValue);

			return boolHolder;
		}

		private byte[] GetDateTimes()
		{
			byte[] data = null;

			List<DateTime> times = [];

			DateTime endUTC = appointmentItem.EndUTC;
			times.Add(endUTC);

			DateTime replyTime = appointmentItem.ReplyTime;
			times.Add(replyTime);

			DateTime startUTC = appointmentItem.StartUTC;
			times.Add(startUTC);

			data = ContentItem.GetDateTimesBytes(times);

			return data;
		}

		private byte[] GetEnums()
		{
			byte[] buffer = null;

			List<int> ints = [];

			int busyStatus = (int)appointmentItem.BusyStatus;
			ints.Add(busyStatus);

			int itemClass = (int)appointmentItem.Class;
			ints.Add(itemClass);

			int importance = (int)appointmentItem.Importance;
			ints.Add(importance);

			int markForDownload = (int)appointmentItem.MarkForDownload;
			ints.Add(markForDownload);

			int meetingStatus = (int)appointmentItem.MeetingStatus;
			ints.Add(meetingStatus);

			int recurrenceState = (int)appointmentItem.RecurrenceState;
			ints.Add(recurrenceState);

			int responseStatus = (int)appointmentItem.ResponseStatus;
			ints.Add(responseStatus);

			int sensitivity = (int)appointmentItem.Sensitivity;
			ints.Add(sensitivity);

			buffer = ContentItem.GetEnumsBuffer(ints);

			return buffer;
		}

		private byte[] GetStringProperties(
			bool strict = false,
			bool ignoreConversation = true)
		{
			byte[] data = null;

			string billingInformation = null;

			try
			{
				billingInformation = appointmentItem.BillingInformation;
			}
			catch (COMException)
			{
			}

			string body = appointmentItem.Body;

			if (body != null && strict == false)
			{
				body = body.TrimEnd();
			}

			string categories = appointmentItem.Categories;
			string companies = appointmentItem.Companies;

			string conversationID = null;

			string conversationTopic = appointmentItem.ConversationTopic;
			string globalAppointmentID = null;

			if (ignoreConversation == false)
			{
				conversationID = appointmentItem.ConversationID;
				globalAppointmentID = appointmentItem.GlobalAppointmentID;
			}

			string location = appointmentItem.Location;
			string meetingWorkspaceURL = appointmentItem.MeetingWorkspaceURL;
			string messageClass = appointmentItem.MessageClass;
			string mileage = appointmentItem.Mileage;
			string optionalAttendees = appointmentItem.OptionalAttendees;
			string organizer = appointmentItem.Organizer;
			string reminderSoundFile = appointmentItem.ReminderSoundFile;
			string requiredAttendees = appointmentItem.RequiredAttendees;
			string resources = appointmentItem.Resources;
			string subject = appointmentItem.Subject;

			StringBuilder builder = new ();
			builder.Append(billingInformation);
			builder.Append(body);
			builder.Append(categories);
			builder.Append(companies);
			builder.Append(conversationID);
			builder.Append(conversationTopic);
			builder.Append(globalAppointmentID);
			builder.Append(location);
			builder.Append(meetingWorkspaceURL);
			builder.Append(messageClass);
			builder.Append(mileage);
			builder.Append(optionalAttendees);
			builder.Append(organizer);
			builder.Append(reminderSoundFile);
			builder.Append(requiredAttendees);
			builder.Append(resources);
			builder.Append(subject);

			string buffer = builder.ToString();

			Encoding encoding = Encoding.UTF8;
			data = encoding.GetBytes(buffer);

			return data;
		}
	}
}
