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
}
