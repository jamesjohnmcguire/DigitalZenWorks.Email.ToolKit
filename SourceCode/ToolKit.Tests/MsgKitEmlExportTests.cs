/////////////////////////////////////////////////////////////////////////////
// <copyright file="MsgKitEmlExportTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.IO;
using System.Threading;
using MimeKit;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Verifies MSG-to-EML preservation for messages saved by Outlook.
/// </summary>
[Apartment(ApartmentState.STA)]
[Category("Integration")]
[NonParallelizable]
internal sealed class MsgKitEmlExportTests
{
	/// <summary>
	/// Verifies headers, both bodies, and a job link survive conversion.
	/// </summary>
	/// <remarks>
	/// Saves a synthetic listing as Unicode MSG, then reopens it in Outlook
	/// to verify the link survived the save. Outlook can recover HTML from
	/// the MSG's compressed RTF body even when MsgKit omits HTML. The test
	/// supplies the recovered HTML and parses the exported EML with MimeKit
	/// to verify that the fallback preserves the listing and its metadata.
	/// </remarks>
	[Test]
	public void ConvertOutlookMsgPreservesJobListing()
	{
		const string listing = "Developer position";
		const string link = "https://example.com/jobs/developer";
		DateTime expectedSent = new(2002, 1, 1, 12, 0, 0, DateTimeKind.Utc);

		using DuplicateTestData data = new();
		string directory = Path.GetDirectoryName(data.StorePath)!;
		string msgPath = Path.Combine(directory, "job-listing.msg");
		string emlPath = Path.Combine(directory, "job-listing.eml");
		Outlook.MailItem item = data.CreateMail();

		try
		{
			item.Subject = listing;
			item.HTMLBody = $"<html><body><a href=\"{link}\">{listing}</a>" +
				"</body></html>";
			Assert.That(item.HTMLBody, Does.Contain(link));
			item.SaveAs(msgPath, Outlook.OlSaveAsType.olMSGUnicode);
		}
		finally
		{
			DuplicateTestData.Release(item);
		}

		// Reopening distinguishes save-time loss from conversion-time loss.
		Outlook.MailItem saved =
			(Outlook.MailItem)data.NameSpace.OpenSharedItem(msgPath);

		string htmlBody;

		try
		{
			htmlBody = saved.HTMLBody ?? string.Empty;
			Assert.That(
				htmlBody,
				Does.Contain(link),
				"HTML link after reopening the saved MSG in Outlook");
		}
		finally
		{
			DuplicateTestData.Release(saved);
		}

		using (FileStream input = File.OpenRead(msgPath))
		using (FileStream output = File.Create(emlPath))
		{
			EmlConverter.ConvertMsgToEml(input, output, htmlBody);
		}

		using MimeMessage exported = MimeMessage.Load(emlPath);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(exported.From.ToString(), Does.Contain(
				"sender@example.com"));
			Assert.That(exported.Date.UtcDateTime, Is.EqualTo(expectedSent));
			Assert.That(exported.Subject, Is.EqualTo(listing));
			Assert.That(
				exported.TextBody ?? string.Empty,
				Does.Contain(listing),
				"Plain-text listing content");
			Assert.That(
				exported.HtmlBody ?? string.Empty,
				Does.Contain(listing).And.Contain(link),
				"HTML listing and link");
		}
	}

	/// <summary>
	/// Verifies a failed move cleans up the staged export.
	/// </summary>
	/// <remarks>
	/// A directory at the final filename blocks the move after conversion.
	/// The failed export should leave only that directory at the destination.
	/// </remarks>
	[Test]
	public void ExportEmlCleansUpWhenDestinationIsBlocked()
	{
		using DuplicateTestData data = new();
		string directory = Path.GetDirectoryName(data.StorePath)!;
		string outputDirectory = Path.Combine(directory, "exports");
		Outlook.MailItem item = data.CreateMail();

		try
		{
			string outputPath = Path.Combine(
				outputDirectory, item.EntryID + ".eml");
			Directory.CreateDirectory(outputPath);
			OutlookMail mail = new(item);

			Assert.That(
				() => mail.ExportEml(outputDirectory),
				Throws.InstanceOf<IOException>().Or
					.InstanceOf<UnauthorizedAccessException>());
			Assert.That(Directory.Exists(outputPath), Is.True);
			Assert.That(
				Directory.GetFileSystemEntries(outputDirectory),
				Is.EqualTo(new[] { outputPath }));
		}
		finally
		{
			DuplicateTestData.Release(item);
		}
	}

	/// <summary>
	/// Verifies the export preserves HTML and skips existing files.
	/// </summary>
	/// <remarks>
	/// Parses the exported EML to verify the source link survived. A second
	/// export targets the same entry ID after sentinel content is placed in
	/// the destination, making an accidental overwrite observable.
	/// </remarks>
	[Test]
	public void ExportEmlWritesHtmlAndPreservesExistingFile()
	{
		const string link = "https://example.com/jobs/developer";
		const string existing = "Existing export";

		using DuplicateTestData data = new();
		string directory = Path.GetDirectoryName(data.StorePath)!;
		string outputDirectory = Path.Combine(directory, "exports");
		Outlook.MailItem item = data.CreateMail();

		try
		{
			item.HTMLBody = $"<html><body><a href=\"{link}\">" +
				"Developer position</a></body></html>";
			string outputPath = Path.Combine(
				outputDirectory, item.EntryID + ".eml");
			OutlookMail mail = new(item);

			bool exported = mail.ExportEml(outputDirectory);
			Assert.That(exported, Is.True);

			using (MimeMessage message = MimeMessage.Load(outputPath))
			{
				Assert.That(
					message.HtmlBody ?? string.Empty,
					Does.Contain(link));
			}

			File.WriteAllText(outputPath, existing);
			bool exportedAgain = mail.ExportEml(outputDirectory);

			Assert.That(exportedAgain, Is.False);
			Assert.That(File.ReadAllText(outputPath), Is.EqualTo(existing));
			Assert.That(
				Directory.GetFiles(outputDirectory),
				Is.EqualTo(new[] { outputPath }));
		}
		finally
		{
			DuplicateTestData.Release(item);
		}
	}

	/// <summary>
	/// Verifies folder export writes only mail from the selected folder.
	/// </summary>
	/// <remarks>
	/// A post in the selected folder and a mail item in a child folder make
	/// type filtering and nonrecursive traversal independently observable.
	/// </remarks>
	[Test]
	public void ExportFolderWritesOnlyMailFromSelectedFolder()
	{
		using DuplicateTestData data = new();
		string directory = Path.GetDirectoryName(data.StorePath)!;
		string outputDirectory = Path.Combine(directory, "exports");
		Outlook.MailItem mail = data.CreateMail();
		Outlook.MAPIFolder folder = data.OpenFolder();
		Outlook.Folders folders = folder.Folders;
		Outlook.Items items = folder.Items;

		try
		{
			Outlook.MAPIFolder child = folders.Add("Child");
			Outlook.MailItem copy = DuplicateTestData.Copy(mail);
			Outlook.PostItem post = (Outlook.PostItem)items.Add(
				Outlook.OlItemType.olPostItem);

			try
			{
				Outlook.MailItem moved = (Outlook.MailItem)copy.Move(child);
				DuplicateTestData.Release(moved);
				post.Subject = "Non-mail fixture";
				post.Save();
			}
			finally
			{
				DuplicateTestData.Release(post, copy, child);
			}

			mail.UnRead = true;
			mail.Save();
			string entryId = mail.EntryID;
			string body = mail.Body;
			string expected = Path.Combine(outputDirectory, entryId + ".eml");
			OutlookStore store = OutlookStore.FromSession(data.Session);

			bool exported = store.ExportEml(
				data.StorePath, "Candidates", outputDirectory);

			using (Assert.EnterMultipleScope())
			{
				Assert.That(exported, Is.True);
				Assert.That(
					Directory.GetFiles(outputDirectory),
					Is.EqualTo(new[] { expected }));
				Assert.That(mail.UnRead, Is.True);
				Assert.That(mail.EntryID, Is.EqualTo(entryId));
				Assert.That(mail.Body, Is.EqualTo(body));
			}
		}
		finally
		{
			DuplicateTestData.Release(items, folders, folder, mail);
		}
	}
}
