/////////////////////////////////////////////////////////////////////////////
// <copyright file="DuplicateCompareTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>Checks meaningful differences and valid hash controls.</summary>
[Apartment(ApartmentState.STA)]
[Category("Integration")]
[NonParallelizable]
internal sealed class DuplicateCompareTests
{
	private static readonly TestCaseData[] BoundaryCases =
	[
		new("AB", "C", "A", "BC"),
		new("A:", "B", "A", ":B"),
		new("A😀", "B", "A", "😀B")
	];

	/// <summary>
	/// Equal-size attachment content differences must remain visible.
	/// </summary>
	[Test]
	public void AttachmentBytesRemainDistinct()
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = DuplicateTestData.Copy(first);

		try
		{
			data.AddDifferentAttachments(first, second);
			Outlook.Attachments firstAttachments = first.Attachments;
			Outlook.Attachments secondAttachments = second.Attachments;
			Outlook.Attachment left = firstAttachments[1];
			Outlook.Attachment right = secondAttachments[1];

			try
			{
				Assert.That(right.DisplayName, Is.EqualTo(left.DisplayName));
				Assert.That(right.FileName, Is.EqualTo(left.FileName));
				Assert.That(right.Index, Is.EqualTo(left.Index));
				Assert.That(right.Position, Is.EqualTo(left.Position));
				Assert.That(right.Type, Is.EqualTo(left.Type));
				string? leftPath = ReadAttachmentPath(left);
				string? rightPath = ReadAttachmentPath(right);
				Assert.That(rightPath, Is.EqualTo(leftPath));
			}
			finally
			{
				DuplicateTestData.Release(
					left, right, firstAttachments, secondAttachments);
			}

			AssertMatchingSynopses(first, second);
			string firstHash = DuplicateTestData.RequiredHash(first);
			string secondHash = DuplicateTestData.RequiredHash(second);
			Assert.That(secondHash, Is.Not.EqualTo(firstHash));
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}
	}

	/// <summary>A meaningful field change must not be a duplicate.</summary>
	/// <param name="field">
	/// The field changed after the equal-copy control.
	/// </param>
	[TestCase("BCC")]
	[TestCase("Body")]
	[TestCase("CC")]
	[TestCase("HTMLBody")]
	[TestCase("Subject")]
	[TestCase("To")]
	[TestCase("UnRead")]
	[TestCase("UserProperty")]
	public void ChangedMailFieldRemainsDistinct(string field)
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = DuplicateTestData.Copy(first);

		try
		{
			string firstHash = DuplicateTestData.RequiredHash(first);
			string copyHash = DuplicateTestData.RequiredHash(second);
			Assert.That(
				copyHash,
				Is.EqualTo(firstHash),
				"Unchanged copies must match before testing the difference.");
			ChangeField(second, field);
			second.Save();

			if (field != "Subject")
			{
				AssertMatchingSynopses(first, second);
			}

			string changedHash = DuplicateTestData.RequiredHash(second);
			Assert.That(changedHash, Is.Not.EqualTo(firstHash));
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}
	}

	/// <summary>
	/// Changed transport headers must distinguish equal content.
	/// </summary>
	[Test]
	public void DifferentHeadersRemainDistinct()
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = data.CreateMail(headerValue: "different");

		try
		{
			Assert.That(second.Subject, Is.EqualTo(first.Subject));
			Assert.That(second.Body, Is.EqualTo(first.Body));
			Assert.That(second.HTMLBody, Is.EqualTo(first.HTMLBody));
			AssertMatchingSynopses(first, second);
			string firstHash = DuplicateTestData.RequiredHash(first);
			string secondHash = DuplicateTestData.RequiredHash(second);
			Assert.That(secondHash, Is.Not.EqualTo(firstHash));
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}
	}

	/// <summary>Equal copies require successful, nonempty hashes.</summary>
	[Test]
	public void EqualCopiesHaveNonemptyHashes()
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = DuplicateTestData.Copy(first);

		try
		{
			string firstHash = DuplicateTestData.RequiredHash(first);
			string secondHash = DuplicateTestData.RequiredHash(second);
			Assert.That(secondHash, Is.EqualTo(firstHash));
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}
	}

	/// <summary>
	/// Repeated content sent at a later date is a separate message.
	/// </summary>
	[Test]
	public void SeparateSendTimesRemainDistinct()
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = data.CreateMail(true);

		try
		{
			Assert.That(second.Subject, Is.EqualTo(first.Subject));
			Assert.That(second.Body, Is.EqualTo(first.Body));
			Assert.That(second.SentOn, Is.Not.EqualTo(first.SentOn));
			string firstHash = DuplicateTestData.RequiredHash(first);
			string secondHash = DuplicateTestData.RequiredHash(second);
			Assert.That(secondHash, Is.Not.EqualTo(firstHash));
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}
	}

	/// <summary>
	/// Distinct field boundaries must remain visible in hashes.
	/// </summary>
	/// <param name="firstCategories">The first category text.</param>
	/// <param name="firstCompanies">The first company text.</param>
	/// <param name="secondCategories">The changed category text.</param>
	/// <param name="secondCompanies">The changed company text.</param>
	[TestCaseSource(nameof(BoundaryCases))]
	public void StringFieldBoundariesRemainDistinct(
		string firstCategories,
		string firstCompanies,
		string secondCategories,
		string secondCompanies)
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = DuplicateTestData.Copy(first);

		try
		{
			DuplicateTestData.SetBoundaryFields(
				first, firstCategories, firstCompanies);
			DuplicateTestData.SetBoundaryFields(
				second, firstCategories, firstCompanies);
			string firstHash = DuplicateTestData.RequiredHash(first);
			string controlHash = DuplicateTestData.RequiredHash(second);

			Assert.That(controlHash, Is.EqualTo(firstHash));
			DuplicateTestData.SetBoundaryFields(
				second, secondCategories, secondCompanies);
			Assert.That(first.Categories, Is.EqualTo(firstCategories));
			Assert.That(first.Companies, Is.EqualTo(firstCompanies));
			AssertMatchingSynopses(first, second);
			string changedHash = DuplicateTestData.RequiredHash(second);
			Assert.That(
				changedHash,
				Is.Not.EqualTo(firstHash),
				"Distinct mail fields must not collapse into the same hash.");
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}
	}

	/// <summary>
	/// Both hash entry points must return the same valid result.
	/// </summary>
	/// <returns>The asynchronous test.</returns>
	[Test]
	public async Task SyncAndAsyncHashesAgree()
	{
		using DuplicateTestData data = new();
		Outlook.MailItem item = data.CreateMail();

		try
		{
			string expected = DuplicateTestData.RequiredHash(item);
			string? actual = await OutlookItem.GetHashAsync(item)
				.ConfigureAwait(true);
			Assert.That(actual, Is.Not.Null.And.Not.Empty);
			Assert.That(actual, Is.EqualTo(expected));
		}
		finally
		{
			DuplicateTestData.Release(item);
		}
	}

	private static void AssertMatchingSynopses(
		Outlook.MailItem first, Outlook.MailItem second)
	{
		string firstSynopsis = OutlookMail.GetSynopses(first);
		string secondSynopsis = OutlookMail.GetSynopses(second);
		Assert.That(
			secondSynopsis,
			Is.EqualTo(firstSynopsis),
			"The synopsis guard alone must not distinguish this case.");
	}

	private static void ChangeField(Outlook.MailItem item, string field)
	{
		switch (field)
		{
			case "BCC":
				item.BCC = "bcc@example.com";
				break;
			case "Body":
				item.Body = "Different body.";
				break;
			case "CC":
				item.CC = "cc@example.com";
				break;
			case "HTMLBody":
				item.HTMLBody = "<html><body>Different HTML.</body></html>";
				break;
			case "Subject":
				item.Subject = "Different subject";
				break;
			case "To":
				item.To = "other@example.com";
				break;
			case "UnRead":
				item.UnRead = true;
				break;
			case "UserProperty":
				DuplicateTestData.SetUserProperty(item, "Unique value");
				break;
			default:
				throw new ArgumentOutOfRangeException(nameof(field));
		}
	}

	private static string? ReadAttachmentPath(Outlook.Attachment attachment)
	{
		try
		{
			return attachment.PathName;
		}
		catch (COMException)
		{
			return null;
		}
	}
}
