/////////////////////////////////////////////////////////////////////////////
// <copyright file="DuplicateTestData.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>Owns synthetic messages and an isolated temporary PST.</summary>
internal sealed class DuplicateTestData : IDisposable
{
	/// <summary>The transport-header property selected by the hash.</summary>
	public const string HeaderProperty =
		"http://schemas.microsoft.com/mapi/proptag/0x007D001F";

	private readonly OutlookTestContext context = new();
	private readonly HashSet<string> attachmentCaches = [];

	/// <summary>
	/// Initializes a new instance of the <see cref="DuplicateTestData"/>
	/// class with a private candidate folder in a test-owned PST.
	/// </summary>
	public DuplicateTestData()
	{
		Session = new(context.NameSpace);
		StorePath = context.StorePath("duplicates.pst");
		Outlook.Store store = Session.GetStore(StorePath)!;
		Outlook.MAPIFolder root = store.GetRootFolder();
		Outlook.Folders folders = root.Folders;

		try
		{
			StoreId = store.StoreID;
			Outlook.MAPIFolder folder = folders.Add("Candidates");
			FolderId = folder.EntryID;
			Release(folder);
		}
		finally
		{
			Release(folders, root, store);
		}
	}

	/// <summary>Gets the candidate folder ID.</summary>
	public string FolderId { get; }

	/// <summary>
	/// Gets the real namespace for direct outcome inspection.
	/// </summary>
	public Outlook.NameSpace NameSpace => context.NameSpace;

	/// <summary>Gets the session wrapping that namespace.</summary>
	public OutlookSession Session { get; }

	/// <summary>Gets the test store ID.</summary>
	public string StoreId { get; }

	/// <summary>Gets the test-owned PST path.</summary>
	public string StorePath { get; }

	/// <summary>Creates a saved copy in the same folder.</summary>
	/// <param name="original">The source mail item.</param>
	/// <returns>The independent copy.</returns>
	public static Outlook.MailItem Copy(Outlook.MailItem original)
	{
		Outlook.MailItem copy = (Outlook.MailItem)original.Copy();
		copy.Save();

		return copy;
	}

	/// <summary>Releases references acquired by the test.</summary>
	/// <param name="references">References still owned by the test.</param>
	public static void Release(params object[] references)
	{
		foreach (object reference in references)
		{
			Marshal.ReleaseComObject(reference);
		}
	}

	/// <summary>Checks hash validity before a comparison can pass.</summary>
	/// <param name="item">The mail item.</param>
	/// <returns>A nonempty production hash.</returns>
	public static string RequiredHash(Outlook.MailItem item)
	{
		OutlookItem wrapper = new(item);
		string? hash = wrapper.Hash;
		Assert.That(
			hash,
			Is.Not.Null.And.Not.Empty,
			"A failed hash must not make a comparison pass.");

		return hash!;
	}

	/// <summary>
	/// Sets distinct mail fields without relying on hash output.
	/// </summary>
	/// <param name="item">The item to update.</param>
	/// <param name="categories">The category text.</param>
	/// <param name="companies">The company text.</param>
	public static void SetBoundaryFields(
		Outlook.MailItem item, string categories, string companies)
	{
		item.Categories = categories;
		item.Companies = companies;
		item.Save();
		Assert.That(item.Categories, Is.EqualTo(categories));
		Assert.That(item.Companies, Is.EqualTo(companies));
		Assert.That(item.CC, Is.Null.Or.Empty);
	}

	/// <summary>Adds a custom value using Outlook's supported API.</summary>
	/// <param name="item">The mail item.</param>
	/// <param name="value">The custom value.</param>
	public static void SetUserProperty(Outlook.MailItem item, string value)
	{
		Outlook.UserProperties properties = item.UserProperties;
		Outlook.UserProperty property = properties.Add(
			"SafetyField", Outlook.OlUserPropertyType.olText, false);

		try
		{
			property.Value = value;
			item.Save();
			Assert.That((string)property.Value, Is.EqualTo(value));
		}
		finally
		{
			Release(property, properties);
		}
	}

	/// <summary>
	/// Adds equal-size attachments with matching names and metadata.
	/// </summary>
	/// <param name="first">The first item.</param>
	/// <param name="second">The second item.</param>
	public void AddDifferentAttachments(
		Outlook.MailItem first, Outlook.MailItem second)
	{
		string name = "DuplicateSafety." + Guid.NewGuid().ToString("N");
		name += ".bin";
		string cache = Path.Combine(Path.GetTempPath(), name);
		attachmentCaches.Add(cache);
		AddAttachment(first, "first", name, [1, 2, 3, 4]);
		AddAttachment(second, "second", name, [1, 2, 3, 5]);
	}

	/// <summary>
	/// Places a sentinel explicitly in the test's Deleted Items.
	/// </summary>
	/// <returns>The moved, test-owned mail item.</returns>
	public Outlook.MailItem CreateDeletedSentinel()
	{
		Outlook.MailItem item = CreateMail();
		Outlook.MAPIFolder folder = OpenFolder();
		Outlook.Store store = folder.Store;
		Outlook.MAPIFolder deleted = store.GetDefaultFolder(
			Outlook.OlDefaultFolders.olFolderDeletedItems);

		try
		{
			Outlook.MailItem moved = (Outlook.MailItem)item.Move(deleted);
			moved.Save();
			Outlook.MAPIFolder parent = (Outlook.MAPIFolder)moved.Parent;

			try
			{
				Assert.That(parent.EntryID, Is.EqualTo(deleted.EntryID));
			}
			finally
			{
				Release(parent);
			}

			return moved;
		}
		finally
		{
			Release(item, deleted, store, folder);
		}
	}

	/// <summary>
	/// Imports synthetic mail with reproducible sender and send time.
	/// </summary>
	/// <param name="laterDate">Whether to model a separate later send.</param>
	/// <param name="headerValue">The synthetic extension-header value.</param>
	/// <returns>The saved mail item.</returns>
	public Outlook.MailItem CreateMail(
		bool laterDate = false, string headerValue = "original")
	{
		DateTime date = new(2002, 1, 1, 12, 0, 0, DateTimeKind.Utc);

		if (laterDate)
		{
			date = date.AddDays(1);
		}

		string dateText = date.ToString("r", CultureInfo.InvariantCulture);
		string[] lines =
		[
			"From: Safety Sender <sender@example.com>",
			"To: Safety Recipient <recipient@example.com>",
			"Subject: Duplicate safety fixture",
			"Date: " + dateText,
			"Message-ID: <duplicate-safety@example.com>",
			"MIME-Version: 1.0",
			"X-Duplicate-Safety: " + headerValue,
			"Content-Type: text/html; charset=utf-8",
			string.Empty,
			"<html><body>Original body.</body></html>",
			string.Empty
		];
		string contents = string.Join("\r\n", lines);
		string fileName = Guid.NewGuid().ToString("N") + ".eml";
		string path = Path.Combine(context.DirectoryPath, fileName);
		File.WriteAllText(path, contents);
		Outlook.MAPIFolder folder = OpenFolder();

		try
		{
			Outlook.MailItem item = Migrate.EmlFileToPst(path, folder);
			Assert.That(item, Is.Not.Null);
			item.UnRead = false;
			item.Save();

			return item;
		}
		finally
		{
			Release(folder);
		}
	}

	/// <summary>Detaches the test store and cleans test-owned files.</summary>
	public void Dispose()
	{
		try
		{
			context.Dispose();
		}
		finally
		{
			foreach (string path in attachmentCaches)
			{
				File.Delete(path);
			}
		}
	}

	/// <summary>Reads the test store's actual Deleted Items count.</summary>
	/// <returns>The deleted item count.</returns>
	public int GetDeletedCount()
	{
		Outlook.MAPIFolder folder = OpenFolder();
		Outlook.Store store = folder.Store;
		Outlook.MAPIFolder deleted = store.GetDefaultFolder(
			Outlook.OlDefaultFolders.olFolderDeletedItems);
		Outlook.Items items = deleted.Items;

		try
		{
			return items.Count;
		}
		finally
		{
			Release(items, deleted, store, folder);
		}
	}

	/// <summary>
	/// Reads survivors directly, independently of the hash code.
	/// </summary>
	/// <returns>The surviving entry IDs.</returns>
	public HashSet<string> GetItemIds()
	{
		HashSet<string> result = [];
		Outlook.MAPIFolder folder = OpenFolder();
		Outlook.Items items = folder.Items;

		try
		{
			for (int index = 1; index <= items.Count; index++)
			{
				Outlook.MailItem item = (Outlook.MailItem)items[index];

				try
				{
					result.Add(item.EntryID);
				}
				finally
				{
					Release(item);
				}
			}
		}
		finally
		{
			Release(items, folder);
		}

		return result;
	}

	/// <summary>Acquires a fresh reference to the candidate folder.</summary>
	/// <returns>The candidate folder.</returns>
	public Outlook.MAPIFolder OpenFolder()
	{
		Outlook.MAPIFolder folder = NameSpace.GetFolderFromID(
			FolderId, StoreId);

		return folder;
	}

	/// <summary>Exercises either production removal implementation.</summary>
	/// <param name="useAsync">Whether to use the asynchronous path.</param>
	/// <param name="dryRun">Whether deletion is disabled.</param>
	/// <returns>The reported duplicate count.</returns>
	public Task<int> RemoveDuplicatesAsync(bool useAsync, bool dryRun)
	{
		OutlookFolder remover = new(Session);
		Outlook.MAPIFolder folder = OpenFolder();

		// Production traversal releases the folder reference it receives.
		if (useAsync)
		{
			return remover.RemoveDuplicatesAsync(folder, dryRun);
		}

		int count = remover.RemoveDuplicates(folder, dryRun);

		return Task.FromResult(count);
	}

	private void AddAttachment(
		Outlook.MailItem item, string directory, string name, byte[] bytes)
	{
		string path = Path.Combine(context.DirectoryPath, directory);
		Directory.CreateDirectory(path);
		string file = Path.Combine(path, name);
		File.WriteAllBytes(file, bytes);
		Outlook.Attachments attachments = item.Attachments;

		try
		{
			Outlook.Attachment attachment = attachments.Add(
				file, Outlook.OlAttachmentType.olByValue, 1, name);
			Release(attachment);
			item.Save();
		}
		finally
		{
			Release(attachments);
		}
	}
}
