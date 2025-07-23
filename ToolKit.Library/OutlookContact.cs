/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookContact.cs" company="James John McGuire">
// Copyright © 2021 - 2025 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using DigitalZenWorks.Common.Utilities;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// OutlookContact Class.
	/// </summary>
	public class OutlookContact
	{
		private readonly ContactItem contactItem;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookContact"/> class.
		/// </summary>
		/// <param name="mapiItem">The Outlook item.</param>
		public OutlookContact(object mapiItem)
		{
#if NET6_0_OR_GREATER
			ArgumentNullException.ThrowIfNull(mapiItem);
#else
			if (mapiItem == null)
			{
				throw new ArgumentNullException(nameof(mapiItem));
			}
#endif

			contactItem = mapiItem as ContactItem;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <param name="contactItem">The AppointmentItem to check.</param>
		/// <returns>The synoses of the item.</returns>
		public static string GetSynopses(ContactItem contactItem)
		{
			string synopses = null;

			if (contactItem != null)
			{
				string sentOn = contactItem.Birthday.ToString(
					"yyyy-MM-dd", CultureInfo.InvariantCulture);

				synopses = string.Format(
					CultureInfo.InvariantCulture,
					"{0}: From: {1}: {2} Subject: {3}",
					sentOn,
					contactItem.FullName,
					contactItem.Email1Address,
					contactItem.Subject);
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

			byte[] buffer = OutlookItem.GetActions(contactItem.Actions);
			buffers.Add(buffer);

			buffer = OutlookItem.GetAttachments(contactItem.Attachments);
			buffers.Add(buffer);

			buffer = GetDateTimesBytes();
			buffers.Add(buffer);

			buffer = GetEnums();
			buffers.Add(buffer);

			buffer = GetStringPropertiesBytes(strict);
			buffers.Add(buffer);

			buffer = OutlookItem.GetUserProperties(
				contactItem.UserProperties);
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
				OutlookItem.GetActionsText(contactItem.Actions);
			propertiesText += Environment.NewLine;

			propertiesText +=
				OutlookItem.GetAttachmentsText(contactItem.Attachments);
			propertiesText += Environment.NewLine;

			propertiesText += GetDateTimesText();
			propertiesText += Environment.NewLine;

			propertiesText += GetEnumsText();
			propertiesText += Environment.NewLine;

			propertiesText += GetStringProperties(strict);
			propertiesText += Environment.NewLine;

			propertiesText += OutlookItem.GetUserPropertiesText(
				contactItem.UserProperties);
			propertiesText += Environment.NewLine;

			return propertiesText;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <returns>The synoses of the item.</returns>
		public string GetSynopses()
		{
			string synopses = GetSynopses(contactItem);

			return synopses;
		}

		private ushort GetBooleans()
		{
			ushort boolHolder = 0;

			bool rawValue = contactItem.AutoResolvedWinner;
			boolHolder = BitBytes.SetBit(boolHolder, 1, rawValue);

			rawValue = contactItem.HasPicture;
			boolHolder = BitBytes.SetBit(boolHolder, 2, rawValue);

			rawValue = contactItem.IsConflict;
			boolHolder = BitBytes.SetBit(boolHolder, 3, rawValue);

			rawValue = contactItem.IsMarkedAsTask;
			boolHolder = BitBytes.SetBit(boolHolder, 4, rawValue);

			rawValue = contactItem.Journal;
			boolHolder = BitBytes.SetBit(boolHolder, 4, rawValue);

			rawValue = contactItem.NoAging;
			boolHolder = BitBytes.SetBit(boolHolder, 5, rawValue);

			rawValue = contactItem.ReminderOverrideDefault;
			boolHolder = BitBytes.SetBit(boolHolder, 6, rawValue);

			rawValue = contactItem.ReminderPlaySound;
			boolHolder = BitBytes.SetBit(boolHolder, 7, rawValue);

			rawValue = contactItem.ReminderSet;
			boolHolder = BitBytes.SetBit(boolHolder, 8, rawValue);

			rawValue = contactItem.Saved;
			boolHolder = BitBytes.SetBit(boolHolder, 10, rawValue);

			rawValue = contactItem.UnRead;
			boolHolder = BitBytes.SetBit(boolHolder, 11, rawValue);
			return boolHolder;
		}

		private string GetBooleansText()
		{
			string booleansText = string.Empty;

			booleansText += OutlookItem.GetBooleanText(
				contactItem.AutoResolvedWinner);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.HasPicture);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.IsConflict);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.IsMarkedAsTask);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.Journal);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.NoAging);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.ReminderOverrideDefault);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.ReminderPlaySound);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.ReminderSet);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.Saved);
			booleansText += OutlookItem.GetBooleanText(
				contactItem.UnRead);

			return booleansText;
		}

		private ReadOnlyCollection<DateTime> GetDateTimes()
		{
			List<DateTime> times = [];

			DateTime time = contactItem.Birthday;
			times.Add(time);

			time = contactItem.ReminderTime;
			times.Add(time);

			time = contactItem.TaskCompletedDate;
			times.Add(time);

			time = contactItem.TaskDueDate;
			times.Add(time);

			time = contactItem.TaskStartDate;
			times.Add(time);

			time = contactItem.ToDoTaskOrdinal;
			times.Add(time);

			ReadOnlyCollection<DateTime> collection = times.AsReadOnly();

			return collection;
		}

		private byte[] GetDateTimesBytes()
		{
			ReadOnlyCollection<DateTime> times = GetDateTimes();

			byte[] data = OutlookItem.GetDateTimesBytes(times);

			return data;
		}

		private string GetDateTimesText()
		{
			ReadOnlyCollection<DateTime> times = GetDateTimes();

			string dateTimesText = OutlookItem.GetDateTimesText(times);

			return dateTimesText;
		}

		private byte[] GetEnums()
		{
			List<int> ints = [];

			int item = (int)contactItem.BusinessCardType;
			ints.Add(item);

			item = (int)contactItem.Class;
			ints.Add(item);

			item = (int)contactItem.DownloadState;
			ints.Add(item);

			item = (int)contactItem.Gender;
			ints.Add(item);

			item = (int)contactItem.Importance;
			ints.Add(item);

			item = (int)contactItem.MarkForDownload;
			ints.Add(item);

			item = (int)contactItem.SelectedMailingAddress;
			ints.Add(item);

			item = (int)contactItem.Sensitivity;
			ints.Add(item);

			byte[] buffer = OutlookItem.GetEnumsBuffer(ints);

			return buffer;
		}

		private string GetEnumsText()
		{
			string enumsText = string.Empty;

			enumsText += nameof(contactItem.BusinessCardType);
			enumsText += nameof(contactItem.Class);
			enumsText += nameof(contactItem.DownloadState);
			enumsText += nameof(contactItem.Gender);
			enumsText += nameof(contactItem.Importance);
			enumsText += nameof(contactItem.MarkForDownload);
			enumsText += nameof(contactItem.SelectedMailingAddress);
			enumsText += nameof(contactItem.Sensitivity);

			return enumsText;
		}

		private string GetStringProperties(
			bool strict = false,
			bool ignoreConversation = true)
		{
			List<string> buffers = [];

			string buffer = contactItem.Account;
			buffers.Add(buffer);

			buffer = contactItem.AssistantName;
			buffers.Add(buffer);

			buffer = contactItem.AssistantTelephoneNumber;
			buffers.Add(buffer);

			try
			{
				buffer = contactItem.BillingInformation;
			}
			catch (COMException)
			{
			}

			buffers.Add(buffer);

			buffer = contactItem.Body;

			if (buffer != null && strict == false)
			{
				buffer = buffer.TrimEnd();
			}

			buffers.Add(buffer);

			buffer = contactItem.Business2TelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.BusinessAddress;
			buffers.Add(buffer);

			buffer = contactItem.BusinessAddressCity;
			buffers.Add(buffer);

			buffer = contactItem.BusinessAddressCountry;
			buffers.Add(buffer);

			buffer = contactItem.BusinessAddressPostalCode;
			buffers.Add(buffer);

			buffer = contactItem.BusinessAddressPostOfficeBox;
			buffers.Add(buffer);

			buffer = contactItem.BusinessAddressState;
			buffers.Add(buffer);

			buffer = contactItem.BusinessAddressStreet;
			buffers.Add(buffer);

			buffer = contactItem.BusinessCardLayoutXml;
			buffers.Add(buffer);

			buffer = contactItem.BusinessFaxNumber;
			buffers.Add(buffer);

			buffer = contactItem.BusinessHomePage;
			buffers.Add(buffer);

			buffer = contactItem.BusinessTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.CallbackTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.CarTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.Categories;
			buffers.Add(buffer);

			buffer = contactItem.Children;
			buffers.Add(buffer);

			buffer = contactItem.Companies;
			buffers.Add(buffer);

			buffer = contactItem.CompanyAndFullName;
			buffers.Add(buffer);

			buffer = contactItem.CompanyLastFirstNoSpace;
			buffers.Add(buffer);

			buffer = contactItem.CompanyLastFirstSpaceOnly;
			buffers.Add(buffer);

			buffer = contactItem.CompanyMainTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.ComputerNetworkName;
			buffers.Add(buffer);

			buffer = contactItem.ConversationTopic;
			buffers.Add(buffer);

			if (ignoreConversation == false)
			{
				buffer = contactItem.ConversationID;
				buffers.Add(buffer);
			}

			buffer = contactItem.CustomerID;
			buffers.Add(buffer);

			buffer = contactItem.Department;
			buffers.Add(buffer);

			buffer = contactItem.Email1Address;
			buffers.Add(buffer);

			buffer = contactItem.Email1AddressType;
			buffers.Add(buffer);

			buffer = contactItem.Email1DisplayName;
			buffers.Add(buffer);

			buffer = contactItem.Email2Address;
			buffers.Add(buffer);

			buffer = contactItem.Email2AddressType;
			buffers.Add(buffer);

			buffer = contactItem.Email2DisplayName;
			buffers.Add(buffer);

			buffer = contactItem.Email3Address;
			buffers.Add(buffer);

			buffer = contactItem.Email3AddressType;
			buffers.Add(buffer);

			buffer = contactItem.Email3DisplayName;
			buffers.Add(buffer);

			buffer = contactItem.FileAs;
			buffers.Add(buffer);

			buffer = contactItem.FirstName;
			buffers.Add(buffer);

			buffer = contactItem.FTPSite;
			buffers.Add(buffer);

			buffer = contactItem.FullName;
			buffers.Add(buffer);

			buffer = contactItem.FullNameAndCompany;
			buffers.Add(buffer);

			buffer = contactItem.GovernmentIDNumber;
			buffers.Add(buffer);

			buffer = contactItem.Hobby;
			buffers.Add(buffer);

			buffer = contactItem.Home2TelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.HomeAddress;
			buffers.Add(buffer);

			buffer = contactItem.HomeAddressCity;
			buffers.Add(buffer);

			buffer = contactItem.HomeAddressCountry;
			buffers.Add(buffer);

			buffer = contactItem.HomeAddressPostalCode;
			buffers.Add(buffer);

			buffer = contactItem.HomeAddressPostOfficeBox;
			buffers.Add(buffer);

			buffer = contactItem.HomeAddressState;
			buffers.Add(buffer);

			buffer = contactItem.HomeAddressStreet;
			buffers.Add(buffer);

			buffer = contactItem.HomeFaxNumber;
			buffers.Add(buffer);

			buffer = contactItem.HomeTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.IMAddress;
			buffers.Add(buffer);

			buffer = contactItem.Initials;
			buffers.Add(buffer);

			buffer = contactItem.InternetFreeBusyAddress;
			buffers.Add(buffer);

			buffer = contactItem.ISDNNumber;
			buffers.Add(buffer);

			buffer = contactItem.JobTitle;
			buffers.Add(buffer);

			buffer = contactItem.Language;
			buffers.Add(buffer);

			buffer = contactItem.LastFirstAndSuffix;
			buffers.Add(buffer);

			buffer = contactItem.LastFirstNoSpace;
			buffers.Add(buffer);

			buffer = contactItem.LastFirstNoSpaceAndSuffix;
			buffers.Add(buffer);

			buffer = contactItem.LastFirstNoSpaceCompany;
			buffers.Add(buffer);

			buffer = contactItem.LastFirstSpaceOnly;
			buffers.Add(buffer);

			buffer = contactItem.LastFirstSpaceOnlyCompany;
			buffers.Add(buffer);

			buffer = contactItem.LastName;
			buffers.Add(buffer);

			buffer = contactItem.LastNameAndFirstName;
			buffers.Add(buffer);

			buffer = contactItem.MailingAddress;
			buffers.Add(buffer);

			buffer = contactItem.MailingAddressCity;
			buffers.Add(buffer);

			buffer = contactItem.MailingAddressCountry;
			buffers.Add(buffer);

			buffer = contactItem.MailingAddressPostalCode;
			buffers.Add(buffer);

			buffer = contactItem.MailingAddressPostOfficeBox;
			buffers.Add(buffer);

			buffer = contactItem.MailingAddressState;
			buffers.Add(buffer);

			buffer = contactItem.MailingAddressStreet;
			buffers.Add(buffer);

			buffer = contactItem.ManagerName;
			buffers.Add(buffer);

			buffer = contactItem.MessageClass;
			buffers.Add(buffer);

			buffer = contactItem.MiddleName;
			buffers.Add(buffer);

			buffer = contactItem.Mileage;
			buffers.Add(buffer);

			buffer = contactItem.MobileTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.NetMeetingAlias;
			buffers.Add(buffer);

			buffer = contactItem.NetMeetingServer;
			buffers.Add(buffer);

			buffer = contactItem.NickName;
			buffers.Add(buffer);

			buffer = contactItem.OfficeLocation;
			buffers.Add(buffer);

			buffer = contactItem.OrganizationalIDNumber;
			buffers.Add(buffer);

			buffer = contactItem.OtherAddress;
			buffers.Add(buffer);

			buffer = contactItem.OtherAddressCity;
			buffers.Add(buffer);

			buffer = contactItem.OtherAddressCountry;
			buffers.Add(buffer);

			buffer = contactItem.OtherAddressPostalCode;
			buffers.Add(buffer);

			buffer = contactItem.OtherAddressPostOfficeBox;
			buffers.Add(buffer);

			buffer = contactItem.OtherAddressState;
			buffers.Add(buffer);

			buffer = contactItem.OtherAddressStreet;
			buffers.Add(buffer);

			buffer = contactItem.OtherFaxNumber;
			buffers.Add(buffer);

			buffer = contactItem.OtherTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.PagerNumber;
			buffers.Add(buffer);

			buffer = contactItem.PersonalHomePage;
			buffers.Add(buffer);

			buffer = contactItem.PrimaryTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.Profession;
			buffers.Add(buffer);

			buffer = contactItem.RadioTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.ReferredBy;
			buffers.Add(buffer);

			buffer = contactItem.ReminderSoundFile;
			buffers.Add(buffer);

			buffer = contactItem.Spouse;
			buffers.Add(buffer);

			buffer = contactItem.Subject;
			buffers.Add(buffer);

			buffer = contactItem.Suffix;
			buffers.Add(buffer);

			buffer = contactItem.TaskSubject;
			buffers.Add(buffer);

			buffer = contactItem.TelexNumber;
			buffers.Add(buffer);

			buffer = contactItem.Title;
			buffers.Add(buffer);

			buffer = contactItem.TTYTDDTelephoneNumber;
			buffers.Add(buffer);

			buffer = contactItem.User1;
			buffers.Add(buffer);

			buffer = contactItem.User2;
			buffers.Add(buffer);

			buffer = contactItem.User3;
			buffers.Add(buffer);

			buffer = contactItem.User4;
			buffers.Add(buffer);

			buffer = contactItem.WebPage;
			buffers.Add(buffer);

			buffer = contactItem.YomiCompanyName;
			buffers.Add(buffer);

			buffer = contactItem.YomiFirstName;
			buffers.Add(buffer);

			buffer = contactItem.YomiLastName;
			buffers.Add(buffer);

			StringBuilder builder = new ();

			foreach (string item in buffers)
			{
				builder.Append(item);
			}

			string stringBuffer = builder.ToString();

			return stringBuffer;
		}

		private byte[] GetStringPropertiesBytes(
			bool strict = false,
			bool ignoreConversation = true)
		{
			string stringBuffer =
				GetStringProperties(strict, ignoreConversation);

			Encoding encoding = Encoding.UTF8;

			byte[] data = encoding.GetBytes(stringBuffer);

			return data;
		}
	}
}
