/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookSessionIntegrationTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Exercises real namespace operations with isolated temporary PSTs and items.
/// </summary>
[Apartment(ApartmentState.STA)]
[NonParallelizable]
internal sealed class OutlookSessionIntegrationTests
{
	/// <summary>
	/// Verifies both folder lookup paths address the requested store and folder.
	/// </summary>
	[Test]
	public void GetFolderFromIdRetrievesRequestedFolder()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);

		string path = context.StorePath("folders.pst");
		Outlook.Store store = session.GetStore(path)!;
		context.Track(store);

		Outlook.MAPIFolder root = store.GetRootFolder();
		context.Track(root);

		Outlook.Folders folders = root.Folders;
		context.Track(folders);

		Outlook.MAPIFolder original = folders.Add("Lookup");
		context.Track(original);

		Outlook.MAPIFolder? retrieved =
			session.GetFolderFromIdInternal(original.EntryID, store.StoreID);
		context.TrackIfNotNull(retrieved);

		Assert.That(retrieved, Is.Not.Null);
		Assert.That(retrieved!.EntryID, Is.EqualTo(original.EntryID));
		Assert.That(retrieved.StoreID, Is.EqualTo(store.StoreID));

		OutlookFolder? wrapped =
			session.GetFolderFromId(original.EntryID, store.StoreID);
		Assert.That(wrapped, Is.Not.Null);

		OutlookFolder child = wrapped!.AddFolder("Child");
		Assert.That(child, Is.Not.Null);

		Outlook.Folders children = original.Folders;
		context.Track(children);
		Assert.That(children.Count, Is.EqualTo(1));
	}

	/// <summary>
	/// Verifies invalid item identifiers propagate the Outlook error.
	/// </summary>
	[Test]
	public void GetItemFromIdInvalidIdThrowsComException()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);

		Assert.That(
			() => session.GetItemFromId("invalid"),
			Throws.TypeOf<COMException>());
	}

	/// <summary>
	/// Verifies lookup retrieves the saved item's identity and contents.
	/// </summary>
	[Test]
	public void GetItemFromIdRetrievesSavedItem()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);

		string path = context.StorePath("items.pst");
		Outlook.Store store = session.GetStore(path)!;
		context.Track(store);

		Outlook.MAPIFolder root = store.GetRootFolder();
		context.Track(root);

		Outlook.Items items = root.Items;
		context.Track(items);

		Outlook.MailItem original = (Outlook.MailItem)items.Add("IPM.Note");
		context.Track(original);

		original.Subject = "Email.ToolKit lookup " + Guid.NewGuid();
		original.Body = "Saved in a test-owned PST.";
		original.Save();

		object? result = session.GetItemFromId(original.EntryID);

		Assert.That(result, Is.InstanceOf<Outlook.MailItem>());

		Outlook.MailItem retrieved = (Outlook.MailItem)result!;
		context.Track(retrieved);

		Assert.That(retrieved.EntryID, Is.EqualTo(original.EntryID));
		Assert.That(retrieved.Subject, Is.EqualTo(original.Subject));
		Assert.That(retrieved.Body, Does.Contain("test-owned PST"));
	}

	/// <summary>
	/// Verifies an existing detached PST is attached only when creation is
	/// allowed; lookup alone leaves the file detached.
	/// </summary>
	/// <param name="createIfMissing">Whether attachment is allowed.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void GetStoreAttachesExistingPstOnlyWhenRequested(
		bool createIfMissing)
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);

		string path = CreateDetachedPst(context, session, "detached.pst");
		Assert.That(File.Exists(path), Is.True);
		Assert.That(context.CountStores(path), Is.Zero);

		Outlook.Store? store = session.GetStore(path, createIfMissing);
		context.TrackIfNotNull(store);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(File.Exists(path), Is.True);

			if (createIfMissing)
			{
				Assert.That(store, Is.Not.Null);
				Assert.That(context.CountStores(path), Is.EqualTo(1));
			}
			else
			{
				Assert.That(store, Is.Null);
				Assert.That(context.CountStores(path), Is.Zero);
			}
		}
	}

	/// <summary>
	/// Verifies the normalized PST exists and is attached exactly once.
	/// </summary>
	/// <param name="fileName">The path supplied to the library.</param>
	/// <param name="expectedName">The expected PST file name.</param>
	[TestCase("store.pst", "store.pst")]
	[TestCase("store.PST", "store.PST")]
	[TestCase("store", "store.pst")]
	[TestCase("store.data", "store.data.pst")]
	[TestCase("nested/../store.pst", "store.pst")]
	public void GetStoreCreatesExpectedPst(string fileName, string expectedName)
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);

		string expectedPath = context.StorePath(expectedName);
		string suppliedPath = Path.Combine(context.DirectoryPath, fileName);

		Outlook.Store? store = session.GetStore(suppliedPath);
		context.TrackIfNotNull(store);

		Assert.That(store, Is.Not.Null);
		Assert.That(store!.FilePath, Is.EqualTo(expectedPath).IgnoreCase);
		Assert.That(File.Exists(expectedPath), Is.True);
		Assert.That(context.CountStores(expectedPath), Is.EqualTo(1));
	}

	/// <summary>
	/// Verifies a missing PST is created and attached only when requested.
	/// </summary>
	/// <param name="createIfMissing">Whether creation is allowed.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void GetStoreCreatesMissingPstOnlyWhenRequested(bool createIfMissing)
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = context.StorePath("missing.pst");
		Assert.That(File.Exists(path), Is.False);
		Assert.That(context.CountStores(path), Is.Zero);

		Outlook.Store? store = session.GetStore(path, createIfMissing);
		context.TrackIfNotNull(store);

		using (Assert.EnterMultipleScope())
		{
			if (createIfMissing)
			{
				Assert.That(store, Is.Not.Null);
				Assert.That(File.Exists(path), Is.True);
				Assert.That(context.CountStores(path), Is.EqualTo(1));
			}
			else
			{
				Assert.That(store, Is.Null);
				Assert.That(File.Exists(path), Is.False);
				Assert.That(context.CountStores(path), Is.Zero);
			}
		}
	}

	/// <summary>
	/// Verifies lookup preserves returned stores and references already held
	/// by callers while releasing its temporary COM references.
	/// </summary>
	[Test]
	public void GetStoreLookupPreservesCallerOwnedStores()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string firstPath = context.StorePath("first.pst");
		string secondPath = context.StorePath("second.pst");
		Outlook.Store first = session.GetStore(firstPath)!;
		context.Track(first);
		string firstId = first.StoreID;
		Outlook.Store second = session.GetStore(secondPath)!;
		context.Track(second);
		string secondId = second.StoreID;

		Outlook.Store? firstLookup = session.GetStore(firstPath, false);
		context.TrackIfNotNull(firstLookup);

		Assert.That(firstLookup, Is.Not.Null);
		Outlook.Store? secondLookup = session.GetStore(secondPath, false);
		context.TrackIfNotNull(secondLookup);

		Assert.That(secondLookup, Is.Not.Null);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(first.StoreID, Is.EqualTo(firstId));
			Assert.That(second.StoreID, Is.EqualTo(secondId));
			Assert.That(firstLookup!.StoreID, Is.EqualTo(firstId));
			Assert.That(secondLookup!.StoreID, Is.EqualTo(secondId));
			Assert.That(context.CountStores(firstPath), Is.EqualTo(1));
			Assert.That(context.CountStores(secondPath), Is.EqualTo(1));
		}
	}

	/// <summary>
	/// Verifies either creation option finds the attached PST with different
	/// path casing and does not attach a duplicate.
	/// </summary>
	/// <param name="createIfMissing">Whether creation is allowed.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void GetStoreReusesExistingStore(bool createIfMissing)
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = context.StorePath("existing.pst");
		Outlook.Store first = session.GetStore(path)!;
		context.Track(first);
		string expectedId = first.StoreID;
		Assert.That(File.Exists(path), Is.True);
		Assert.That(context.CountStores(path), Is.EqualTo(1));

		string uppercasePath = path.ToUpperInvariant();
		Outlook.Store? second =
			session.GetStore(uppercasePath, createIfMissing);
		context.TrackIfNotNull(second);

		Assert.That(second, Is.Not.Null);
		Assert.That(second!.StoreID, Is.EqualTo(expectedId));
		Assert.That(context.CountStores(path), Is.EqualTo(1));
	}

	/// <summary>
	/// Verifies removal accepts the same extensionless path used to create a PST.
	/// </summary>
	/// <param name="withCollidingFile">Whether a file without the PST suffix
	/// also exists at the supplied path.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void GetStoreThenRemoveStoreAcceptsExtensionlessPath(
		bool withCollidingFile)
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = context.StorePath("round-trip.pst");
		string extensionlessPath = Path.Combine(
			context.DirectoryPath, "round-trip");

		if (withCollidingFile)
		{
			File.WriteAllText(extensionlessPath, "Not a PST file.");
		}

		Outlook.Store? store = session.GetStore(extensionlessPath);
		context.TrackIfNotNull(store);

		Assert.That(store, Is.Not.Null);
		Assert.That(context.CountStores(path), Is.EqualTo(1));

		bool removed = session.RemoveStore(extensionlessPath);

		Assert.That(removed, Is.True);
		Assert.That(context.CountStores(path), Is.Zero);

		if (withCollidingFile)
		{
			Assert.That(
				File.ReadAllText(extensionlessPath),
				Is.EqualTo("Not a PST file."));
		}
	}

	/// <summary>
	/// Verifies a non-mail shared item is rejected by the mail wrapper.
	/// </summary>
	[Test]
	public void OpenMailItemFileRejectsAppointment()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		Outlook.AppointmentItem original = (Outlook.AppointmentItem)
			context.Application.CreateItem(
				Outlook.OlItemType.olAppointmentItem);
		context.Track(original);
		original.Subject = "Email.ToolKit appointment";
		string path = Path.Combine(context.DirectoryPath, "appointment.msg");
		original.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);

		Assert.That(session.OpenMailItemFile(path), Is.Null);
		original.Close(Outlook.OlInspectorClose.olDiscard);
	}

	/// <summary>
	/// Verifies the mail wrapper can read an actual saved mail item.
	/// </summary>
	[Test]
	public void OpenMailItemFileWrapsMail()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = Path.Combine(context.DirectoryPath, "wrapped.msg");
		Outlook.MailItem original = SaveMailItemAsMsg(
			context, path, "Email.ToolKit wrapped mail");

		OutlookMail? opened = session.OpenMailItemFile(path);

		Assert.That(opened, Is.Not.Null);
		Assert.That(opened!.GetSynopses(), Does.Contain(original.Subject));
		original.Close(Outlook.OlInspectorClose.olDiscard);
	}

	/// <summary>
	/// Verifies a missing shared file propagates the Outlook error.
	/// </summary>
	[Test]
	public void OpenSharedItemMissingFileThrowsFileNotFoundException()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = Path.Combine(context.DirectoryPath, "missing.msg");

		Assert.That(
			() => session.OpenSharedItem(path),
			Throws.TypeOf<FileNotFoundException>());
	}

	/// <summary>
	/// Verifies shared mail files retain their known contents.
	/// </summary>
	[Test]
	public void OpenSharedItemReadsSavedMail()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = Path.Combine(context.DirectoryPath, "mail.msg");
		Outlook.MailItem original = SaveMailItemAsMsg(
			context,
			path,
			"Email.ToolKit shared mail",
			"Contents from the saved MSG.");

		object? result = session.OpenSharedItem(path);

		Assert.That(result, Is.InstanceOf<Outlook.MailItem>());
		Outlook.MailItem opened = (Outlook.MailItem)result!;
		context.Track(opened);
		Assert.That(opened.Subject, Is.EqualTo(original.Subject));
		Assert.That(opened.Body, Does.Contain("saved MSG"));
		opened.Close(Outlook.OlInspectorClose.olDiscard);
		original.Close(Outlook.OlInspectorClose.olDiscard);
	}

	/// <summary>
	/// Verifies an existing PST file returns false when it is not attached,
	/// without reattaching or deleting it.
	/// </summary>
	[Test]
	public void RemoveStoreReturnsFalseWhenExistingPstIsNotAttached()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = CreateDetachedPst(context, session, "detached.pst");

		Assert.That(File.Exists(path), Is.True);
		Assert.That(context.CountStores(path), Is.Zero);

		bool removed = session.RemoveStore(path);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(removed, Is.False);
			Assert.That(context.CountStores(path), Is.Zero);
			Assert.That(File.Exists(path), Is.True);
		}
	}

	/// <summary>
	/// Verifies removal returns false without creating or attaching a missing
	/// PST file.
	/// </summary>
	[Test]
	public void RemoveStoreReturnsFalseWhenPstDoesNotExist()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = context.StorePath("missing.pst");

		bool fileExists = File.Exists(path);

		Assert.That(fileExists, Is.False);
		Assert.That(context.CountStores(path), Is.Zero);

		bool removed = session.RemoveStore(path);

		fileExists = File.Exists(path);
		int count = context.CountStores(path);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(removed, Is.False);
			Assert.That(fileExists, Is.False);
			Assert.That(count, Is.Zero);
		}
	}

	/// <summary>
	/// Verifies detaching a PST returns true and removing it again by path
	/// returns false, while retaining its file.
	/// </summary>
	/// <param name="byPath">Whether to call the path overload first.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void RemoveStoreReturnsTrueWhenPstIsDetached(bool byPath)
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = context.StorePath("remove.pst");
		Outlook.Store store = session.GetStore(path)!;
		context.Track(store);
		Assert.That(context.CountStores(path), Is.EqualTo(1));

		bool removed;

		if (byPath)
		{
			removed = session.RemoveStore(path);
		}
		else
		{
			removed = session.RemoveStore(store);
		}

		using (Assert.EnterMultipleScope())
		{
			Assert.That(removed, Is.True);
			Assert.That(context.CountStores(path), Is.Zero);
			Assert.That(File.Exists(path), Is.True);
		}

		bool removedAgain = session.RemoveStore(path);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(removedAgain, Is.False);
			Assert.That(context.CountStores(path), Is.Zero);
			Assert.That(File.Exists(path), Is.True);
		}
	}

	/// <summary>
	/// Verifies removal without a session preserves the caller's store and
	/// root folder references and leaves the store attached.
	/// </summary>
	[Test]
	public void RemoveStoreWithoutSessionPreservesCallerOwnedReferences()
	{
		using OutlookTestContext context = new();
		OutlookSession connectedSession = new(context.NameSpace);
		string path = context.StorePath("retained.pst");
		Outlook.Store store = connectedSession.GetStore(path)!;
		context.Track(store);
		Outlook.MAPIFolder root = store.GetRootFolder();
		context.Track(root);
		string storeId = store.StoreID;
		string rootId = root.EntryID;
		OutlookSession session = new(null);

		bool removed = session.RemoveStore(store);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(removed, Is.False);
			Assert.That(store.StoreID, Is.EqualTo(storeId));
			Assert.That(root.EntryID, Is.EqualTo(rootId));
			Assert.That(context.CountStores(path), Is.EqualTo(1));
		}
	}

	private static string CreateDetachedPst(
		OutlookTestContext context, OutlookSession session, string fileName)
	{
		string path = context.StorePath(fileName);
		Outlook.Store store = session.GetStore(path)!;
		context.Track(store);
		Outlook.MAPIFolder root = store.GetRootFolder();
		context.Track(root);
		context.NameSpace.RemoveStore(root);
		return path;
	}

	private static Outlook.MailItem SaveMailItemAsMsg(
		OutlookTestContext context,
		string path,
		string subject,
		string? body = null)
	{
		Outlook.MailItem item = (Outlook.MailItem)
			context.Application.CreateItem(
				Outlook.OlItemType.olMailItem);
		context.Track(item);
		item.Subject = subject;

		if (body != null)
		{
			item.Body = body;
		}

		item.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);
		return item;
	}
}
