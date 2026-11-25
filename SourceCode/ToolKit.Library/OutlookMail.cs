/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookMail.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using DigitalZenWorks.Common.Utilities;
using Microsoft.Office.Interop.Outlook;

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
#if NET6_0_OR_GREATER
		ArgumentNullException.ThrowIfNull(mapiItem);
#else
		if (mapiItem == null)
		{
			throw new ArgumentNullException(nameof(mapiItem));
		}
#endif

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
	/// Exports this mail item's text and available HTML bodies.
	/// </summary>
	/// <param name="outputDirectory">The directory for entry-ID-named files.
	/// </param>
	/// <returns>True if files were written; false if output already exists or
	/// this wrapper does not contain a mail item.</returns>
	/// <remarks>This experimental export preserves the body text supplied by
	/// Outlook. File-system and Outlook errors propagate to the caller.
	/// </remarks>
	public bool ExportBodies(string outputDirectory)
	{
		bool result = false;

		if (mailItem != null)
		{
			string sentOn = mailItem.SentOn.ToString(
				"O", CultureInfo.InvariantCulture);
			string sender = GetSenderLine();

			result = WriteBodiesWrapper(
				mailItem.EntryID,
				sentOn,
				sender,
				mailItem.Body,
				mailItem.HTMLBody,
				outputDirectory);
		}

		return result;
	}

	/// <summary>
	/// Get the bytes of all relevant properties.
	/// </summary>
	/// <param name="strict">Indicates whether the check should be strict
	/// or not.</param>
	/// <returns>The bytes of all relevant properties.</returns>
	/// <remarks>Recipients is already included with
	/// string properties.</remarks>
	public IList<byte[]> GetProperties(bool strict = false)
	{
		List<byte[]> buffers = [];

		ushort booleans = GetBooleans();

		byte[] buffer = OutlookItem.GetActions(mailItem.Actions);
		buffers.Add(buffer);

		buffer = OutlookItem.GetAttachments(mailItem.Attachments);
		buffers.Add(buffer);

		buffer = GetDateTimesBytes();
		buffers.Add(buffer);

		buffer = GetEnums();
		buffers.Add(buffer);

		buffer = GetStringProperties(strict);
		buffers.Add(buffer);

		buffer = OutlookItem.GetUserProperties(
			mailItem.UserProperties);
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
	/// <param name="addNewLine">Indicates whether to add a new line
	/// or not.</param>
	/// <returns>The text of all relevant properties.</returns>
	public string GetPropertiesText(
		bool strict = false, bool addNewLine = false)
	{
		string propertiesText = string.Empty;

		propertiesText += GetBooleansText();
		propertiesText += Environment.NewLine;

		propertiesText +=
			OutlookItem.GetActionsText(mailItem.Actions, addNewLine);
		propertiesText += Environment.NewLine;

		propertiesText +=
			OutlookItem.GetAttachmentsText(mailItem.Attachments);
		propertiesText += Environment.NewLine;

		propertiesText += GetDateTimesText();
		propertiesText += Environment.NewLine;

		propertiesText += GetEnumsText();
		propertiesText += Environment.NewLine;

		propertiesText +=
			GetStringPropertiesText(strict, true, addNewLine);
		propertiesText += Environment.NewLine;

		propertiesText += OutlookItem.GetUserPropertiesText(
			mailItem.UserProperties);
		propertiesText += Environment.NewLine;

		return propertiesText;
	}

	/// <summary>
	/// Get the item's synopses.
	/// </summary>
	/// <returns>The synopses of the item.</returns>
	public string GetSynopses()
	{
		string synopses = GetSynopses(mailItem);

		return synopses;
	}

	/// <summary>
	/// Writes text and available HTML unless either export already exists.
	/// </summary>
	/// <param name="entryId">The entry ID used for output filenames.</param>
	/// <param name="sentOn">The formatted sent date and time.</param>
	/// <param name="sender">The sender metadata.</param>
	/// <param name="text">The plain-text body.</param>
	/// <param name="html">The available HTML body.</param>
	/// <param name="directory">The output directory.</param>
	/// <returns>True if files were written; false if output exists.</returns>
	internal static bool WriteBodiesWrapper(
		string entryId,
		string sentOn,
		string sender,
		string text,
		string html,
		string directory)
	{
		bool result = false;

		string textPath = Path.Combine(directory, entryId + ".txt");
		string htmlPath = Path.Combine(directory, entryId + ".html");
		bool textExists = File.Exists(textPath);
		bool htmlExists = File.Exists(htmlPath);

		if (textExists == false && htmlExists == false)
		{
			Directory.CreateDirectory(directory);
			WriteFile(textPath, sentOn, sender, text, false);

			bool exists = !string.IsNullOrEmpty(html);

			if (string.IsNullOrEmpty(html) == false)
			{
				WriteFile(htmlPath, sentOn, sender, html, true);
			}

			result = true;
		}

		return result;
	}

	/// <summary>
	/// Writes a UTF-8 body file with sent-time and sender metadata.
	/// </summary>
	/// <param name="path">The new output file path.</param>
	/// <param name="sentOn">The formatted sent date and time.</param>
	/// <param name="sender">The sender metadata.</param>
	/// <param name="body">The body content to preserve.</param>
	/// <param name="html">Whether to write metadata as an HTML comment.</param>
	internal static void WriteFile(
		string path, string sentOn, string sender, string body, bool html)
	{
		// Keep sender metadata on one line in both output formats.
		sender = sender.Replace('\r', ' ');
		sender = sender.Replace('\n', ' ');

		using FileStream stream =
			new(path, FileMode.CreateNew, FileAccess.Write);
		UTF8Encoding encoding = new(false);
		using StreamWriter writer = new(stream, encoding);
		writer.NewLine = "\n";

		if (html == true)
		{
			sender = WebUtility.HtmlEncode(sender);
			StringBuilder escapedSender = new(sender);
			escapedSender.Replace("--", "&#45;&#45;");
			sender = escapedSender.ToString();
			writer.WriteLine("<!-- Sent: " + sentOn);
			writer.WriteLine("Sender: " + sender + " -->");
		}
		else
		{
			writer.WriteLine("Sent: " + sentOn);
			writer.WriteLine("Sender: " + sender);
		}

		writer.WriteLine();
		writer.Write(body);
	}

	/// <summary>
	/// Exports the Outlook item to an entry-ID-named EML file.
	/// </summary>
	/// <param name="outputDirectory">The destination directory.</param>
	/// <returns>True if an EML was written, or false if the item has no entry
	/// ID or the output file already exists.</returns>
	/// <remarks>Outlook and file-system errors propagate to the caller.
	/// </remarks>
	internal bool ExportEml(string outputDirectory)
	{
		bool exported = false;
		string entryId = mailItem?.EntryID;

		if (string.IsNullOrEmpty(entryId) == false)
		{
			string outputPath = Path.Combine(outputDirectory, entryId + ".eml");

			if (File.Exists(outputPath) == false)
			{
				Directory.CreateDirectory(outputDirectory);
				WriteEml(outputDirectory, outputPath);
				exported = true;
			}
		}

		return exported;
	}

	private static string GetTextBoundaryMarker(
		string text, bool preserveBoundaries)
	{
		string boundaryMarker = string.Empty;

		if (preserveBoundaries == true)
		{
			int length = 0;

			if (text != null)
			{
				length = text.Length;
			}

			boundaryMarker = length.ToString(CultureInfo.InvariantCulture);
			boundaryMarker += ":";
		}

		return boundaryMarker;
	}

	private static string NormalizeHeaders(string headers)
	{
#if NETCOREAPP1_0_OR_GREATER
		string[] parts = headers.Split("\r\n");
#else
		string[] parts = headers.Split('\n');
#endif

		List<string> list = [.. parts];

		list.Sort();

#if NETCOREAPP1_0_OR_GREATER
		headers = string.Join("\r\n", list);
#else
		headers = string.Join("\n", list);
#endif

		return headers;
	}

	private void WriteEml(string outputDirectory, string outputPath)
	{
		string temporaryName = Guid.NewGuid().ToString("N");
		string msgPath = Path.Combine(
			Path.GetTempPath(), temporaryName + ".msg");
		string stagingPath = Path.Combine(
			outputDirectory, temporaryName + ".tmp");

		try
		{
			string htmlBody = mailItem.HTMLBody;
			mailItem.SaveAs(msgPath, OlSaveAsType.olMSGUnicode);

			using (FileStream input = File.OpenRead(msgPath))
			using (FileStream output = new(stagingPath, FileMode.CreateNew))
			{
				EmlConverter.ConvertMsgToEml(input, output, htmlBody);
			}

			// Publish the complete file without replacing an existing export.
			File.Move(stagingPath, outputPath);
		}
		finally
		{
			try
			{
				File.Delete(stagingPath);
			}
			finally
			{
				File.Delete(msgPath);
			}
		}
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

		// IsConflict ?
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

	private string GetBooleansText()
	{
		string booleansText = string.Empty;

		string boolValue = mailItem.AutoForwarded.ToString();
		booleansText +=
			OutlookItem.FormatValue("AutoForwarded", boolValue);

		boolValue = mailItem.AutoResolvedWinner.ToString();
		booleansText +=
			OutlookItem.FormatValue("AutoResolvedWinner", boolValue);

		boolValue = mailItem.DeleteAfterSubmit.ToString();
		booleansText +=
			OutlookItem.FormatValue("DeleteAfterSubmit", boolValue);

		// mailItem.IsConflict
		boolValue = mailItem.IsMarkedAsTask.ToString();
		booleansText +=
			OutlookItem.FormatValue("IsMarkedAsTask", boolValue);

		boolValue = mailItem.NoAging.ToString();
		booleansText +=
			OutlookItem.FormatValue("NoAging", boolValue);

		boolValue = mailItem.OriginatorDeliveryReportRequested.ToString();
		booleansText += OutlookItem.FormatValue(
			"OriginatorDeliveryReportRequested", boolValue);

		boolValue = mailItem.ReadReceiptRequested.ToString();
		booleansText +=
			OutlookItem.FormatValue("ReadReceiptRequested", boolValue);

		boolValue = mailItem.RecipientReassignmentProhibited.ToString();
		booleansText += OutlookItem.FormatValue(
			"RecipientReassignmentProhibited", boolValue);

		boolValue = mailItem.ReminderOverrideDefault.ToString();
		booleansText +=
			OutlookItem.FormatValue("ReminderOverrideDefault", boolValue);

		boolValue = mailItem.ReminderPlaySound.ToString();
		booleansText +=
			OutlookItem.FormatValue("ReminderPlaySound", boolValue);

		boolValue = mailItem.ReminderSet.ToString();
		booleansText +=
			OutlookItem.FormatValue("ReminderSet", boolValue);

		boolValue = mailItem.Saved.ToString();
		booleansText +=
			OutlookItem.FormatValue("Saved", boolValue);

		boolValue = mailItem.Submitted.ToString();
		booleansText +=
			OutlookItem.FormatValue("Submitted", boolValue);

		boolValue = mailItem.UnRead.ToString();
		booleansText +=
			OutlookItem.FormatValue("UnRead", boolValue);

		return booleansText;
	}

	private ReadOnlyCollection<DateTime> GetDateTimes()
	{
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

		List<string> labelsRaw =
		[
			"DeferredDeliveryTimeDateTime",
			"ExpiryTime",
			"ReceivedTime",
			"ReminderTime",
			"RetentionExpirationDate",
			"SentOn",
			"TaskCompletedDate",
			"TaskDueDate",
			"mailItem.TaskStartDate"
		];

		ReadOnlyCollection<string> labels = new(labelsRaw);

		string dateTimesText =
			OutlookItem.GetDateTimesText(times, labels);

		return dateTimesText;
	}

	private byte[] GetEnums()
	{
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

		byte[] buffer = OutlookItem.GetEnumsBuffer(ints);

		return buffer;
	}

	private string GetEnumsText()
	{
		string enumsText = string.Empty;

		enumsText +=
			OutlookItem.FormatEnumValue("BodyFormat", mailItem.BodyFormat);

		enumsText +=
			OutlookItem.FormatEnumValue("Class", mailItem.Class);

		enumsText +=
			OutlookItem.FormatEnumValue("Importance", mailItem.Importance);

		enumsText += OutlookItem.FormatEnumValue(
			"MarkForDownload", mailItem.MarkForDownload);

		enumsText +=
			OutlookItem.FormatEnumValue("Permission", mailItem.Permission);

		enumsText += OutlookItem.FormatEnumValue(
			"PermissionService", mailItem.PermissionService);

		enumsText += OutlookItem.FormatEnumValue(
			"Sensitivity", mailItem.Sensitivity);

		return enumsText;
	}

	private string GetSenderLine()
	{
		string senderLine = null;

		if (mailItem != null)
		{
			string name = mailItem.SenderName ?? string.Empty;
			string address = mailItem.SenderEmailAddress ?? string.Empty;
			senderLine = name;

			bool exists = !string.IsNullOrEmpty(address);

			if (exists == true)
			{
				exists = !string.IsNullOrEmpty(name);

				if (exists == false)
				{
					senderLine = address;
				}
				else
				{
					senderLine = name + " <" + address + ">";
				}
			}
		}

		return senderLine;
	}

	private byte[] GetStringProperties(
		bool strict = false,
		bool ignoreConversation = true)
	{
		string propertiesText =
			GetStringPropertiesText(
				strict, ignoreConversation, preserveBoundaries: true);

		Encoding encoding = Encoding.UTF8;
		byte[] data = encoding.GetBytes(propertiesText);

		return data;
	}

	private string GetStringPropertiesText(
		bool strict = false,
		bool ignoreConversation = true,
		bool formatText = false,
		bool preserveBoundaries = false)
	{
		List<string> properties = [];

		string formattedText = string.Empty;
		string bcc = mailItem.BCC;

		bcc = OutlookItem.FormatValueConditional("BCC", bcc, formatText);
		properties.Add(bcc);

		string billingInformation = null;

		try
		{
			billingInformation = mailItem.BillingInformation;
		}
		catch (COMException)
		{
		}

		billingInformation = OutlookItem.FormatValueConditional(
			"BillingInformation", billingInformation, formatText);
		properties.Add(billingInformation);

		string body = mailItem.Body;

		if (body != null && strict == false)
		{
			body = body.TrimEnd();
		}

		body = OutlookItem.FormatValueConditional(
			"Body", body, formatText);
		properties.Add(body);

		string categories = mailItem.Categories;
		categories = OutlookItem.FormatValueConditional(
			"Categories", categories, formatText);
		properties.Add(categories);

		string cc = mailItem.CC;
		cc = OutlookItem.FormatValueConditional(
			"CC", cc, formatText);
		properties.Add(cc);

		string companies = mailItem.Companies;
		companies = OutlookItem.FormatValueConditional(
			"Companies", companies, formatText);
		properties.Add(companies);

		string conversationID = null;

		if (ignoreConversation == false)
		{
			conversationID = mailItem.ConversationID;
			conversationID = OutlookItem.FormatValueConditional(
				"ConversationID", conversationID, formatText);
			properties.Add(conversationID);
		}

		string conversationTopic = mailItem.ConversationTopic;
		conversationTopic = OutlookItem.FormatValueConditional(
			"ConversationTopic", conversationTopic, formatText);
		properties.Add(conversationTopic);

		string flagRequest = mailItem.FlagRequest;
		flagRequest = OutlookItem.FormatValueConditional(
			"FlagRequest", flagRequest, formatText);
		properties.Add(flagRequest);

		string header = mailItem.PropertyAccessor.GetProperty(
			"http://schemas.microsoft.com/mapi/proptag/0x007D001F");

		if (header != null && strict == false)
		{
			header = OutlookItem.RemoveMimeOleVersion(header);

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

		header = OutlookItem.FormatValueConditional(
			"Header (0x007D001F)", header, formatText);
		properties.Add(header);

		string htmlBody = mailItem.HTMLBody;

		if (htmlBody != null && strict == false)
		{
			htmlBody = HtmlEmail.Trim(htmlBody);
		}

		htmlBody = OutlookItem.FormatValueConditional(
			"HTMLBody", htmlBody, formatText);
		properties.Add(htmlBody);

		string messageClass = mailItem.MessageClass;
		messageClass = OutlookItem.FormatValueConditional(
			"MessageClass", messageClass, formatText);
		properties.Add(messageClass);

		string mileage = mailItem.Mileage;
		mileage = OutlookItem.FormatValueConditional(
			"Mileage", mileage, formatText);
		properties.Add(mileage);

		string receivedByName = mailItem.ReceivedByName;
		receivedByName = OutlookItem.FormatValueConditional(
			"ReceivedByName", receivedByName, formatText);
		properties.Add(receivedByName);

		string reminderSoundFile = mailItem.ReminderSoundFile;
		reminderSoundFile = OutlookItem.FormatValueConditional(
			"ReminderSoundFile", reminderSoundFile, formatText);
		properties.Add(reminderSoundFile);

		string replyRecipientNames = mailItem.ReplyRecipientNames;
		replyRecipientNames = OutlookItem.FormatValueConditional(
			"ReplyRecipientNames", replyRecipientNames, formatText);
		properties.Add(replyRecipientNames);

		string retentionPolicyName = mailItem.RetentionPolicyName;
		retentionPolicyName = OutlookItem.FormatValueConditional(
			"RetentionPolicyName", retentionPolicyName, formatText);
		properties.Add(retentionPolicyName);

		string senderEmailAddress = mailItem.SenderEmailAddress;
		senderEmailAddress = OutlookItem.FormatValueConditional(
			"SenderEmailAddress", senderEmailAddress, formatText);
		properties.Add(senderEmailAddress);

		string senderEmailType = mailItem.SenderEmailType;
		senderEmailType = OutlookItem.FormatValueConditional(
			"SenderEmailType", senderEmailType, formatText);
		properties.Add(senderEmailType);

		string senderName = mailItem.SenderName;
		senderName = OutlookItem.FormatValueConditional(
			"SenderName", senderName, formatText);
		properties.Add(senderName);

		string sentOnBehalfOfName = mailItem.SentOnBehalfOfName;
		sentOnBehalfOfName = OutlookItem.FormatValueConditional(
			"SentOnBehalfOfName", sentOnBehalfOfName, formatText);
		properties.Add(sentOnBehalfOfName);

		string subject = mailItem.Subject;
		subject = OutlookItem.FormatValueConditional(
			"Subject", subject, formatText);
		properties.Add(subject);

		string taskSubject = mailItem.TaskSubject;
		taskSubject = OutlookItem.FormatValueConditional(
			"TaskSubject", taskSubject, formatText);
		properties.Add(taskSubject);

		string to = mailItem.To;
		to = OutlookItem.FormatValueConditional(
			"To", to, formatText);
		properties.Add(to);

		string votingOptions = mailItem.VotingOptions;
		votingOptions = OutlookItem.FormatValueConditional(
			"VotingOptions", votingOptions, formatText);
		properties.Add(votingOptions);

		string votingResponse = mailItem.VotingResponse;
		votingResponse = OutlookItem.FormatValueConditional(
			"VotingResponse", votingResponse, formatText);
		properties.Add(votingResponse);

		if (strict == true)
		{
			// Might need to investigate further.
			string receivedByEntryID = mailItem.ReceivedByEntryID;
			receivedByEntryID = OutlookItem.FormatValueConditional(
				"ReceivedByEntryID", receivedByEntryID, formatText);
			properties.Add(receivedByEntryID);

			string receivedOnBehalfOfEntryID =
				mailItem.ReceivedOnBehalfOfEntryID;
			receivedOnBehalfOfEntryID =
				OutlookItem.FormatValueConditional(
					"ReceivedOnBehalfOfEntryID",
					receivedOnBehalfOfEntryID,
					formatText);
			properties.Add(receivedOnBehalfOfEntryID);

			string receivedOnBehalfOfName =
				mailItem.ReceivedOnBehalfOfName;
			receivedOnBehalfOfName = OutlookItem.FormatValueConditional(
				"ReceivedOnBehalfOfName",
				receivedOnBehalfOfName,
				formatText);
			properties.Add(receivedOnBehalfOfName);
		}

		StringBuilder builder = new();

		foreach (string item in properties)
		{
			string boundaryMarker =
				GetTextBoundaryMarker(item, preserveBoundaries);

			builder.Append(boundaryMarker);
			builder.Append(item);
		}

		string stringProperties = builder.ToString();

		return stringProperties;
	}
}
