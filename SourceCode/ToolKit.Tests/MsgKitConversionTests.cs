/////////////////////////////////////////////////////////////////////////////
// <copyright file="MsgKitConversionTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using MimeKit;
using MsgKit;
using NUnit.Framework;
using OpenMcdf;

/// <summary>
/// Verifies legacy MIME conversion and notification data without Outlook.
/// </summary>
internal sealed class MsgKitConversionTests
{
	private const string AttachmentHeaders =
		"Content-Type: application/octet-stream; name=\"legacy.bin\"\r\n" +
		"Content-Transfer-Encoding: base64\r\n";

	private const string HtmlCid = "<html><body>" +
		"<img src=\"cid:legacy@example.test\"></body></html>";

	private const string HtmlLocation = "<html><body>" +
		"<img src=\"legacy.bin\"></body></html>";

	private const string NotificationFields =
		"Reporting-UA: legacy.example.test; Mail 1998\r\n" +
		"Final-Recipient: rfc822; recipient@example.test\r\n" +
		"Original-Message-ID: <original@example.test>\r\n" +
		"Disposition: manual-action/MDN-sent-manually; displayed\r\n";

	private static readonly byte[] Payload = [0x01, 0x02, 0x03, 0xff];

	/// <summary>
	/// Verifies an attached message is retained without rewriting its parts.
	/// </summary>
	[Test]
	public void EmbeddedMessageRetainsReadReceiptFields()
	{
		byte[] embedded = CreateReadReceipt(string.Empty);
		string attachedMessage = Encoding.ASCII.GetString(embedded);
		byte[] eml = CreateEml(
			"text/plain",
			"An attached receipt.",
			"Content-Type: message/rfc822",
			attachedMessage);

		byte[] msg = ConvertToMsg(eml);

		byte[]? attached = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		Assert.That(attached, Is.Not.Null.And.Not.Empty);
		string stored = Encoding.ASCII.GetString(attached!);
		Assert.That(stored, Does.Contain("recipient@example.test"));
		Assert.That(stored, Does.Contain("displayed"));
	}

	/// <summary>
	/// Verifies matching inline identifiers and HTML survive conversion.
	/// </summary>
	[Test]
	public void MatchingCidPreservesReferenceAndBytes()
	{
		string headers = AttachmentHeaders +
			"Content-Disposition: inline\r\n" +
			"Content-ID: <legacy@example.test>";
		byte[] eml = CreateEml(
			"text/html", HtmlCid, headers, "AQID/w==", "multipart/related");

		byte[] msg = ConvertToMsg(eml);

		byte[]? data = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		Assert.That(data, Is.EqualTo(Payload));
		byte[]? identifier = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_CONTENT_ID_W.Name);
		string contentId = DecodeUnicode(identifier);
		Assert.That(contentId, Is.EqualTo("legacy@example.test"));
		string html = ReadHtml(msg);
		Assert.That(html, Does.Contain(HtmlCid));
	}

	/// <summary>
	/// Verifies absent or unusable identifiers do not discard attachments.
	/// </summary>
	/// <param name="extraHeaders">Optional disposition and ID headers.</param>
	[TestCase("")]
	[TestCase("Content-Disposition: inline")]
	[TestCase("Content-Disposition: attachment")]
	[TestCase("Content-Disposition: inline\r\nContent-ID: \t")]
	[TestCase("Content-Disposition: inline\r\nContent-ID: <>")]
	[TestCase("Content-Disposition: inline\r\nContent-ID: < >")]
	public void MissingOrBlankContentIdPreservesAttachmentBytes(
		string extraHeaders)
	{
		string headers = AttachmentHeaders + extraHeaders;
		byte[] eml = CreateEml(
			"text/plain", "Legacy message body.", headers, "AQID/w==");

		byte[] msg = ConvertToMsg(eml);

		byte[]? data = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		Assert.That(data, Is.EqualTo(Payload));
		byte[]? name = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_LONG_FILENAME_W.Name);
		string fileName = DecodeUnicode(name);
		Assert.That(fileName, Is.EqualTo("legacy.bin"));
		byte[]? contentId = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_CONTENT_ID_W.Name);
		Assert.That(contentId, Is.Null);
		byte[]? body = ReadMessage(msg, PropertyTags.PR_BODY_W.Name);
		string text = DecodeUnicode(body);
		Assert.That(text, Does.Contain("Legacy message body."));
	}

	/// <summary>
	/// Verifies separate receipts cannot overwrite each other's fields.
	/// </summary>
	[Test]
	public void MultipleReadReceiptsRemainSeparate()
	{
		string secondFields = NotificationFields.Replace(
			"recipient@example.test",
			"second@example.test",
			StringComparison.Ordinal);
		string nested = string.Join(
			"\r\n",
			"--nested",
			"Content-Type: message/disposition-notification",
			string.Empty,
			NotificationFields,
			"--nested",
			"Content-Type: message/disposition-notification",
			string.Empty,
			secondFields,
			"--nested--");
		byte[] eml = CreateEml(
			"text/plain",
			"Two read receipts.",
			"Content-Type: multipart/mixed; boundary=\"nested\"",
			nested);

		byte[] msg = ConvertToMsg(eml);

		byte[]? first = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		byte[]? second = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name, 1);
		Assert.That(first, Is.Not.Null.And.Not.Empty);
		Assert.That(second, Is.Not.Null.And.Not.Empty);
		string firstText = Encoding.ASCII.GetString(first!);
		string secondText = Encoding.ASCII.GetString(second!);
		Assert.That(firstText, Does.Contain("recipient@example.test"));
		Assert.That(secondText, Does.Contain("second@example.test"));
		byte[]? firstName = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_LONG_FILENAME_W.Name);
		byte[]? secondName = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_LONG_FILENAME_W.Name, 1);
		string firstFileName = DecodeUnicode(firstName);
		string secondFileName = DecodeUnicode(secondName);
		Assert.That(secondFileName, Is.Not.EqualTo(firstFileName));
	}

	/// <summary>
	/// Verifies the file-based conversion preserves the archived EML bytes.
	/// </summary>
	[Test]
	public void ReadReceiptConversionLeavesSourceFileUnchanged()
	{
		Guid identifier = Guid.NewGuid();
		string name = "EmailToolKit-receipt-" + identifier.ToString("N");
		string directory = Path.Combine(Path.GetTempPath(), name);
		Directory.CreateDirectory(directory);
		string source = Path.Combine(directory, "receipt.eml");
		string destination = Path.Combine(directory, "receipt.msg");
		byte[] original = CreateReadReceipt(string.Empty);

		try
		{
			File.WriteAllBytes(source, original);

			EmlConverter.ConvertEmlToMsg(source, destination);

			byte[] unchanged = File.ReadAllBytes(source);
			Assert.That(unchanged, Is.EqualTo(original));
			byte[] msg = File.ReadAllBytes(destination);
			AssertReadReceiptFields(msg);
		}
		finally
		{
			File.Delete(source);
			File.Delete(destination);
			Directory.Delete(directory);
		}
	}

	/// <summary>
	/// Verifies notifications nested in a multipart retain their fields.
	/// </summary>
	[Test]
	public void ReadReceiptInNestedMultipartPreservesFields()
	{
		string nested = string.Join(
			"\r\n",
			"--nested",
			"Content-Type: message/disposition-notification",
			string.Empty,
			NotificationFields,
			"--nested--");
		byte[] eml = CreateEml(
			"text/plain",
			"Your message was displayed.",
			"Content-Type: multipart/mixed; boundary=\"nested\"",
			nested);
		VerifyReadReceiptInput(eml);

		byte[] msg = ConvertToMsg(eml);

		AssertReadReceiptFields(msg);
	}

	/// <summary>
	/// Verifies notification repair retains an existing name and content ID.
	/// </summary>
	[Test]
	public void ReadReceiptPreservesNameAndContentId()
	{
		string headers =
			"Content-Disposition: inline; filename=\"receipt.txt\"\r\n" +
			"Content-ID: <receipt@example.test>";
		byte[] eml = CreateReadReceipt(headers);

		byte[] msg = ConvertToMsg(eml);

		AssertReadReceiptFields(msg);
		byte[]? name = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_LONG_FILENAME_W.Name);
		byte[]? identifier = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_CONTENT_ID_W.Name);
		string fileName = DecodeUnicode(name);
		string contentId = DecodeUnicode(identifier);
		Assert.That(fileName, Is.EqualTo("receipt.txt"));
		Assert.That(contentId, Is.EqualTo("receipt@example.test"));
	}

	/// <summary>
	/// Verifies read-receipt fields remain available in the MSG attachment.
	/// </summary>
	/// <param name="extraHeaders">Optional notification disposition.</param>
	[TestCase("")]
	[TestCase("Content-Disposition: inline")]
	public void ReadReceiptPreservesNotificationFields(string extraHeaders)
	{
		byte[] eml = CreateReadReceipt(extraHeaders);
		VerifyReadReceiptInput(eml);

		byte[] msg = ConvertToMsg(eml);

		AssertReadReceiptFields(msg);
		byte[]? body = ReadMessage(msg, PropertyTags.PR_BODY_W.Name);
		string text = DecodeUnicode(body);
		Assert.That(text, Does.Contain("Your message was displayed."));
	}

	/// <summary>
	/// Verifies relative-location examples retain their bytes and HTML.
	/// This does not assert how Outlook renders the converted message.
	/// </summary>
	[Test]
	public void RelativeContentLocationRetainsAttachmentBytes()
	{
		string headers = AttachmentHeaders +
			"Content-Disposition: inline\r\nContent-Location: legacy.bin";
		byte[] eml = CreateEml(
			"text/html",
			HtmlLocation,
			headers,
			"AQID/w==",
			"multipart/related");

		byte[] msg = ConvertToMsg(eml);

		byte[]? data = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		Assert.That(data, Is.EqualTo(Payload));
		byte[]? name = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_LONG_FILENAME_W.Name);
		string fileName = DecodeUnicode(name);
		Assert.That(fileName, Is.EqualTo("legacy.bin"));
		string html = ReadHtml(msg);
		Assert.That(html, Does.Contain(HtmlLocation));
	}

	/// <summary>
	/// Verifies a standalone notification MIME body retains its fields.
	/// </summary>
	[Test]
	public void SinglePartNotificationPreservesFields()
	{
		string source = string.Join(
			"\r\n",
			"From: recipient@example.test",
			"To: sender@example.test",
			"Subject: Standalone notification",
			"MIME-Version: 1.0",
			"Content-Type: message/disposition-notification",
			string.Empty,
			NotificationFields);
		byte[] eml = Encoding.ASCII.GetBytes(source);
		VerifyReadReceiptInput(eml);

		byte[] msg = ConvertToMsg(eml);

		AssertReadReceiptFields(msg);
	}

	/// <summary>
	/// Verifies conversion does not invent a target for unmatched HTML IDs.
	/// </summary>
	[Test]
	public void UnmatchedCidRetainsReferenceWithoutInventingId()
	{
		string headers = AttachmentHeaders + "Content-Disposition: inline";
		byte[] eml = CreateEml(
			"text/html", HtmlCid, headers, "AQID/w==", "multipart/related");

		byte[] msg = ConvertToMsg(eml);

		byte[]? data = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		Assert.That(data, Is.EqualTo(Payload));
		byte[]? contentId = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_CONTENT_ID_W.Name);
		Assert.That(contentId, Is.Null);
		string html = ReadHtml(msg);
		Assert.That(html, Does.Contain(HtmlCid));
	}

	/// <summary>
	/// Verifies unnamed MIME attachments retain their decoded bytes.
	/// </summary>
	[Test]
	public void UnnamedAttachmentPreservesBytes()
	{
		string headers = "Content-Type: application/octet-stream\r\n" +
			"Content-Transfer-Encoding: base64\r\n" +
			"Content-Disposition: inline";
		byte[] eml = CreateEml(
			"text/plain", "Legacy message body.", headers, "AQID/w==");

		byte[] msg = ConvertToMsg(eml);

		byte[]? data = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		Assert.That(data, Is.EqualTo(Payload));
		byte[]? name = ReadAttachment(
			msg, PropertyTags.PR_ATTACH_LONG_FILENAME_W.Name);
		string fileName = DecodeUnicode(name);
		Assert.That(fileName, Is.Not.Empty);
	}

	private static void AssertReadReceiptFields(byte[] message)
	{
		byte[]? data = ReadAttachment(
			message, PropertyTags.PR_ATTACH_DATA_BIN.Name);
		Assert.That(
			data,
			Is.Not.Null.And.Not.Empty,
			"The machine-readable read receipt must not be discarded.");
		string stored = Encoding.ASCII.GetString(data!);
		Assert.That(stored, Does.Contain("legacy.example.test"));
		Assert.That(stored, Does.Contain("recipient@example.test"));
		Assert.That(stored, Does.Contain("<original@example.test>"));
		Assert.That(
			stored,
			Does.Contain("manual-action/MDN-sent-manually; displayed"));
	}

	private static byte[] ConvertToMsg(byte[] eml)
	{
		using MemoryStream input = new(eml);
		using MemoryStream output = new();
		EmlConverter.ConvertEmlToMsg(input, output);
		byte[] message = output.ToArray();
		return message;
	}

	private static byte[] CreateEml(
		string bodyType,
		string body,
		string attachmentHeaders,
		string attachmentBody,
		string outerType = "multipart/mixed")
	{
		string headers = attachmentHeaders.TrimEnd('\r', '\n');
		string eml = string.Join(
			"\r\n",
			"From: recipient@example.test",
			"To: sender@example.test",
			"Date: Thu, 01 Jan 1998 12:00:00 +0000",
			"Subject: Legacy MIME regression",
			"MIME-Version: 1.0",
			"Content-Type: " + outerType + "; boundary=\"legacy-boundary\"",
			string.Empty,
			"--legacy-boundary",
			"Content-Type: " + bodyType,
			string.Empty,
			body,
			"--legacy-boundary",
			headers,
			string.Empty,
			attachmentBody,
			"--legacy-boundary--",
			string.Empty);
		byte[] encoded = Encoding.ASCII.GetBytes(eml);
		return encoded;
	}

	private static byte[] CreateReadReceipt(string extraHeaders)
	{
		string headers =
			"Content-Type: message/disposition-notification\r\n" +
			extraHeaders;
		byte[] message = CreateEml(
			"text/plain",
			"Your message was displayed.",
			headers,
			NotificationFields,
			"multipart/report; report-type=disposition-notification");
		return message;
	}

	private static string DecodeUnicode(byte[]? data)
	{
		string text = string.Empty;

		if (data != null)
		{
			text = Encoding.Unicode.GetString(data);
			text = text.TrimEnd('\0');
		}

		return text;
	}

	private static byte[]? ReadAttachment(
		byte[] message, string property, int attachmentIndex = 0)
	{
		using MemoryStream input = new(message);
		using RootStorage root = RootStorage.Open(
			input, StorageModeFlags.LeaveOpen);
		string suffix = attachmentIndex.ToString(
			"X8", CultureInfo.InvariantCulture);
		string name = "__attach_version1.0_#" + suffix;
		Storage attachment = root.OpenStorage(name);
		byte[]? data = ReadProperty(attachment, property);
		return data;
	}

	private static string ReadHtml(byte[] message)
	{
		byte[]? data = ReadMessage(message, PropertyTags.PR_HTML.Name);
		Assert.That(data, Is.Not.Null);
		string html = Encoding.UTF8.GetString(data!);
		return html;
	}

	private static byte[]? ReadMessage(byte[] message, string property)
	{
		using MemoryStream input = new(message);
		using RootStorage root = RootStorage.Open(
			input, StorageModeFlags.LeaveOpen);
		byte[]? data = ReadProperty(root, property);
		return data;
	}

	private static byte[]? ReadProperty(Storage storage, string property)
	{
		CfbStream? stream = null;

		try
		{
			bool found = storage.TryOpenStream(property, out stream);

			if (found == false || stream == null)
			{
				return null;
			}

			using MemoryStream copy = new();
			stream.CopyTo(copy);
			byte[] data = copy.ToArray();
			return data;
		}
		finally
		{
			stream?.Dispose();
		}
	}

	private static void VerifyReadReceiptInput(byte[] message)
	{
		using MemoryStream input = new(message);
		using MimeMessage parsed = MimeMessage.Load(input);
		MimeEntity notification = parsed.BodyParts.Last();
		Assert.That(
			notification,
			Is.InstanceOf<MessageDispositionNotification>());
		using MemoryStream serialized = new();
		notification.WriteTo(serialized);
		byte[] data = serialized.ToArray();
		string fields = Encoding.ASCII.GetString(data);
		Assert.That(fields, Does.Contain("recipient@example.test"));
		Assert.That(
			fields,
			Does.Contain("manual-action/MDN-sent-manually; displayed"));
	}
}
