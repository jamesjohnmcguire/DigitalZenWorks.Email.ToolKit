﻿/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookAppointment.cs" company="James John McGuire">
// Copyright © 2021 - 2025 James John McGuire. All Rights Reserved.
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
	/// OutlookAppointment Class.
	/// </summary>
	public class OutlookAppointment
	{
		private readonly AppointmentItem appointmentItem;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookAppointment"/> class.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		public OutlookAppointment(object mapiItem)
		{
#if NET6_0_OR_GREATER
			ArgumentNullException.ThrowIfNull(mapiItem);
#else
			if (mapiItem == null)
			{
				throw new ArgumentNullException(nameof(mapiItem));
			}
#endif

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
		public IList<byte[]> GetProperties(bool strict = false)
		{
			List<byte[]> buffers = [];

			ushort booleans = GetBooleans();

			byte[] buffer = OutlookItem.GetActions(appointmentItem.Actions);
			buffers.Add(buffer);

			buffer = OutlookItem.GetAttachments(appointmentItem.Attachments);
			buffers.Add(buffer);

			buffer = GetDateTimesBytes();
			buffers.Add(buffer);

			buffer = GetEnums();
			buffers.Add(buffer);

			buffer = OutlookItem.GetRecipients(appointmentItem.Recipients);
			buffers.Add(buffer);

			buffer = GetStringPropertiesBytes(strict);
			buffers.Add(buffer);

			buffer = OutlookItem.GetUserProperties(
				appointmentItem.UserProperties);
			buffers.Add(buffer);

			byte[] itemBytes = new byte[2];
			itemBytes = BitBytes.CopyUshortToByteArray(
				itemBytes, 0, booleans);
			buffers.Add(itemBytes);

			return buffers;
		}

		/// <summary>
		/// Get the text of all relevant properties.
		/// </summary>
		/// <param name="strict">Indicates whether the check should be strict
		/// or not.</param>
		/// <returns>The text of all relevant properties.</returns>
		public string GetPropertiesText(bool strict = false)
		{
			string propertiesText = string.Empty;

			propertiesText += GetBooleansText();
			propertiesText += Environment.NewLine;

			propertiesText +=
				OutlookItem.GetActionsText(appointmentItem.Actions);
			propertiesText += Environment.NewLine;

			propertiesText +=
				OutlookItem.GetAttachmentsText(appointmentItem.Attachments);
			propertiesText += Environment.NewLine;

			propertiesText += GetDateTimesText();
			propertiesText += Environment.NewLine;

			propertiesText += GetEnumsText();
			propertiesText += Environment.NewLine;

			propertiesText +=
				OutlookItem.GetRecipientsText(appointmentItem.Recipients);
			propertiesText += Environment.NewLine;

			propertiesText += GetStringProperties(strict);
			propertiesText += Environment.NewLine;

			propertiesText += OutlookItem.GetUserPropertiesText(
				appointmentItem.UserProperties);
			propertiesText += Environment.NewLine;

			return propertiesText;
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

		private string GetBooleansText()
		{
			string booleansText = string.Empty;

			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.AllDayEvent);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.AutoResolvedWinner);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.ForceUpdateToAllAttendees);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.IsConflict);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.IsRecurring);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.NoAging);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.ReminderOverrideDefault);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.ReminderPlaySound);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.ReminderSet);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.ResponseRequested);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.Saved);
			booleansText += OutlookItem.GetBooleanText(
				appointmentItem.UnRead);

			return booleansText;
		}

		private List<DateTime> GetDateTimes()
		{
			List<DateTime> times = [];

			DateTime endUTC = appointmentItem.EndUTC;
			times.Add(endUTC);

			DateTime replyTime = appointmentItem.ReplyTime;
			times.Add(replyTime);

			DateTime startUTC = appointmentItem.StartUTC;
			times.Add(startUTC);

			return times;
		}

		private byte[] GetDateTimesBytes()
		{
			List<DateTime> times = GetDateTimes();

			byte[] data = OutlookItem.GetDateTimesBytes(times);

			return data;
		}

		private string GetDateTimesText()
		{
			List<DateTime> times = GetDateTimes();

			string dateTimesText = OutlookItem.GetDateTimesText(times);

			return dateTimesText;
		}

		private byte[] GetEnums()
		{
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

			byte[] buffer = OutlookItem.GetEnumsBuffer(ints);

			return buffer;
		}

		private string GetEnumsText()
		{
			string enumsText = string.Empty;

			enumsText += nameof(appointmentItem.BusyStatus);
			enumsText += nameof(appointmentItem.Class);
			enumsText += nameof(appointmentItem.Importance);
			enumsText += nameof(appointmentItem.MarkForDownload);
			enumsText += nameof(appointmentItem.MeetingStatus);
			enumsText += nameof(appointmentItem.RecurrenceState);
			enumsText += nameof(appointmentItem.ResponseStatus);
			enumsText += nameof(appointmentItem.Sensitivity);

			return enumsText;
		}

		private string GetStringProperties(
			bool strict = false,
			bool ignoreConversation = true)
		{
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

			string stringProperties = builder.ToString();

			return stringProperties;
		}

		private byte[] GetStringPropertiesBytes(
			bool strict = false,
			bool ignoreConversation = true)
		{
			string buffer = GetStringProperties(strict, ignoreConversation);

			Encoding encoding = Encoding.UTF8;
			byte[] data = encoding.GetBytes(buffer);

			return data;
		}
	}
}
