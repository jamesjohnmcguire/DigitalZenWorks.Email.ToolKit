/////////////////////////////////////////////////////////////////////////////
// <copyright file="EmlConverter.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System.IO;
using MimeKit;
using MsgKit;

/// <summary>
/// Preserves notification data when converting EML through upstream MsgKit.
/// </summary>
internal static class EmlConverter
{
	/// <summary>
	/// Converts a source EML file without modifying its archived contents.
	/// </summary>
	/// <param name="emlFileName">The source EML file.</param>
	/// <param name="msgFileName">The new MSG file to create.</param>
	internal static void ConvertEmlToMsg(
		string emlFileName, string msgFileName)
	{
		using FileStream input = new(
			emlFileName, FileMode.Open, FileAccess.Read);
		using FileStream output = new(
			msgFileName, FileMode.CreateNew, FileAccess.Write);
		ConvertEmlToMsg(input, output);
	}

	/// <summary>
	/// Converts EML data, preserving read receipts as ordinary attachments.
	/// Messages without notifications are passed to MsgKit unchanged.
	/// </summary>
	/// <param name="emlFile">The readable, seekable EML input stream.</param>
	/// <param name="msgFile">The writable, seekable MSG output stream.</param>
	internal static void ConvertEmlToMsg(Stream emlFile, Stream msgFile)
	{
		long inputPosition = emlFile.Position;
		byte[]? conversionCopy = CreateConversionCopy(emlFile);

		if (conversionCopy == null)
		{
			emlFile.Position = inputPosition;
			Converter.ConvertEmlToMsg(emlFile, msgFile);
		}
		else
		{
			using MemoryStream normalized = new(conversionCopy, false);
			Converter.ConvertEmlToMsg(normalized, msgFile);
		}
	}

	/// <summary>
	/// Serializes a conversion copy only when a notification needs repair.
	/// Disposes the parsed message before MsgKit parses the selected input.
	/// </summary>
	/// <param name="emlFile">The source MIME stream to inspect.</param>
	/// <returns>Normalized EML bytes, or null when no repair is needed.
	/// </returns>
	private static byte[]? CreateConversionCopy(Stream emlFile)
	{
		using MimeMessage message = MimeMessage.Load(emlFile);
		bool normalized = NormalizeReadReceipts(message);

		if (normalized == false)
		{
			return null;
		}

		using MemoryStream copy = new();
		message.WriteTo(copy);
		byte[] data = copy.ToArray();
		return data;
	}

	/// <summary>
	/// Stores the notification body as a normal attachment so MsgKit cannot
	/// discard it in its specialized notification conversion branch.
	/// </summary>
	/// <param name="notification">The parsed notification to preserve.</param>
	/// <returns>An attachment that owns its serialized notification data.
	/// </returns>
	private static MimePart CreateReadReceiptAttachment(
		MessageDispositionNotification notification)
	{
		using MemoryStream serialized = new();
		notification.WriteTo(FormatOptions.Default, serialized, true);
		byte[] fields = serialized.ToArray();
		string? fileName = notification.FileName;

		if (string.IsNullOrWhiteSpace(fileName))
		{
			fileName = "read-receipt.txt";
		}

		MimePart attachment = new("application", "octet-stream");
		attachment.ContentDisposition =
			new ContentDisposition(ContentDisposition.Attachment);
		attachment.FileName = fileName;
		attachment.ContentId = notification.ContentId;
		attachment.ContentLocation = notification.ContentLocation;
		attachment.ContentDescription = notification.ContentDescription;
		attachment.ContentTransferEncoding = ContentEncoding.Base64;
		MemoryStream content = new(fields, false);
		attachment.Content = new MimeContent(content);
		return attachment;
	}

	/// <summary>
	/// Replaces notifications in the MIME body of the conversion copy.
	/// </summary>
	/// <param name="message">The parsed source message.</param>
	/// <returns>True when at least one notification was replaced.</returns>
	private static bool NormalizeReadReceipts(MimeMessage message)
	{
		bool normalized = false;

		if (message.Body is MessageDispositionNotification notification)
		{
			MimePart attachment = CreateReadReceiptAttachment(notification);
			message.Body = attachment;
			notification.Dispose();
			normalized = true;
		}
		else if (message.Body is Multipart multipart)
		{
			normalized = NormalizeReadReceipts(multipart);
		}

		return normalized;
	}

	/// <summary>
	/// Replaces notification parts while retaining attached messages intact.
	/// MsgKit saves attached messages as EML rather than converting their body.
	/// </summary>
	/// <param name="multipart">The multipart body to examine.</param>
	/// <returns>True when at least one notification was replaced.</returns>
	private static bool NormalizeReadReceipts(Multipart multipart)
	{
		bool normalized = false;

		for (int index = 0; index < multipart.Count; index++)
		{
			MimeEntity part = multipart[index];

			if (part is MessageDispositionNotification notification)
			{
				MimePart attachment = CreateReadReceiptAttachment(notification);
				multipart[index] = attachment;
				notification.Dispose();
				normalized = true;
			}
			else if (part is Multipart nested)
			{
				bool nestedChanged = NormalizeReadReceipts(nested);

				if (nestedChanged == true)
				{
					normalized = true;
				}
			}
		}

		return normalized;
	}
}
