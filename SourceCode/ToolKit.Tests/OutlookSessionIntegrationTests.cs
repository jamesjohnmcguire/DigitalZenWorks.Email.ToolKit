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

		Assert.That(store, Is.Not.Null);
		context.Track(store!);
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

		if (store != null)
		{
			context.Track(store);
		}

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
		Outlook.Store first = context.Track(session.GetStore(path) !);
		string expectedId = first.StoreID;
		Assert.That(File.Exists(path), Is.True);
		Assert.That(context.CountStores(path), Is.EqualTo(1));

		Outlook.Store? second =
			session.GetStore(path.ToUpperInvariant(), createIfMissing);

		if (second != null)
		{
			context.Track(second);
		}

		Assert.That(second, Is.Not.Null);
		Assert.That(second!.StoreID, Is.EqualTo(expectedId));
		Assert.That(context.CountStores(path), Is.EqualTo(1));
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
		string path = context.StorePath("detached.pst");
		Outlook.Store original = context.Track(session.GetStore(path) !);
		Outlook.MAPIFolder root = context.Track(original.GetRootFolder());
		context.NameSpace.RemoveStore(root);
		Assert.That(File.Exists(path), Is.True);
		Assert.That(context.CountStores(path), Is.Zero);

		Outlook.Store? store = session.GetStore(path, createIfMissing);

		if (store != null)
		{
			context.Track(store);
		}

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
		Outlook.Store first = context.Track(session.GetStore(firstPath) !);
		string firstId = first.StoreID;
		Outlook.Store second = context.Track(session.GetStore(secondPath) !);
		string secondId = second.StoreID;

		Outlook.Store? firstLookup = session.GetStore(firstPath, false);

		if (firstLookup != null)
		{
			context.Track(firstLookup);
		}

		Assert.That(firstLookup, Is.Not.Null);
		Outlook.Store? secondLookup = session.GetStore(secondPath, false);

		if (secondLookup != null)
		{
			context.Track(secondLookup);
		}

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
	/// Verifies removal without a session preserves the caller's store and
	/// root folder references and leaves the store attached.
	/// </summary>
	[Test]
	public void RemoveStoreWithoutSessionPreservesCallerOwnedReferences()
	{
		using OutlookTestContext context = new();
		OutlookSession connectedSession = new(context.NameSpace);
		string path = context.StorePath("retained.pst");
		Outlook.Store store = context.Track(connectedSession.GetStore(path) !);
		Outlook.MAPIFolder root = context.Track(store.GetRootFolder());
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
		Outlook.Store store = context.Track(session.GetStore(path) !);
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
	/// Verifies an existing PST file returns false when it is not attached,
	/// without reattaching or deleting it.
	/// </summary>
	[Test]
	public void RemoveStoreReturnsFalseWhenExistingPstIsNotAttached()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = context.StorePath("detached.pst");
		Outlook.Store store = context.Track(session.GetStore(path) !);
		Outlook.MAPIFolder root = context.Track(store.GetRootFolder());
		context.NameSpace.RemoveStore(root);

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
	/// Verifies lookup retrieves the saved item's identity and contents.
	/// </summary>
	[Test]
	public void GetItemFromIdRetrievesSavedItem()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		string path = context.StorePath("items.pst");
		Outlook.Store store = context.Track(session.GetStore(path) !);
		Outlook.MAPIFolder root = context.Track(store.GetRootFolder());
		Outlook.Items items = context.Track(root.Items);
		Outlook.MailItem original =
			context.Track((Outlook.MailItem)items.Add("IPM.Note"));
		original.Subject = "Email.ToolKit lookup " + Guid.NewGuid();
		original.Body = "Saved in a test-owned PST.";
		original.Save();

		object? result = session.GetItemFromId(original.EntryID);

		Assert.That(result, Is.InstanceOf<Outlook.MailItem>());
		Outlook.MailItem retrieved =
			context.Track((Outlook.MailItem)result!);
		Assert.That(retrieved.EntryID, Is.EqualTo(original.EntryID));
		Assert.That(retrieved.Subject, Is.EqualTo(original.Subject));
		Assert.That(retrieved.Body, Does.Contain("test-owned PST"));
	}

	/// <summary>
	/// Verifies both folder lookup paths address the requested store and folder.
	/// </summary>
	[Test]
	public void GetFolderFromIdRetrievesRequestedFolder()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		Outlook.Store store = context.Track(
			session.GetStore(context.StorePath("folders.pst")) !);
		Outlook.MAPIFolder root = context.Track(store.GetRootFolder());
		Outlook.Folders folders = context.Track(root.Folders);
		Outlook.MAPIFolder original = context.Track(folders.Add("Lookup"));

		Outlook.MAPIFolder? retrieved =
			session.GetFolderFromIdInternal(original.EntryID, store.StoreID);

		Assert.That(retrieved, Is.Not.Null);
		context.Track(retrieved!);
		Assert.That(retrieved!.EntryID, Is.EqualTo(original.EntryID));
		Assert.That(retrieved.StoreID, Is.EqualTo(store.StoreID));

		OutlookFolder? wrapped =
			session.GetFolderFromId(original.EntryID, store.StoreID);
		Assert.That(wrapped, Is.Not.Null);
		OutlookFolder child = wrapped!.AddFolder("Child");
		Assert.That(child, Is.Not.Null);
		Outlook.Folders children = context.Track(original.Folders);
		Assert.That(children.Count, Is.EqualTo(1));
	}

	/// <summary>
	/// Verifies shared mail files retain their known contents.
	/// </summary>
	[Test]
	public void OpenSharedItemReadsSavedMail()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		Outlook.MailItem original = context.Track(
			(Outlook.MailItem)context.Application.CreateItem(
				Outlook.OlItemType.olMailItem));
		original.Subject = "Email.ToolKit shared mail";
		original.Body = "Contents from the saved MSG.";
		string path = Path.Combine(context.DirectoryPath, "mail.msg");
		original.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);

		object? result = session.OpenSharedItem(path);

		Assert.That(result, Is.InstanceOf<Outlook.MailItem>());
		Outlook.MailItem opened = context.Track((Outlook.MailItem)result!);
		Assert.That(opened.Subject, Is.EqualTo(original.Subject));
		Assert.That(opened.Body, Does.Contain("saved MSG"));
		opened.Close(Outlook.OlInspectorClose.olDiscard);
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
		Outlook.MailItem original = context.Track(
			(Outlook.MailItem)context.Application.CreateItem(
				Outlook.OlItemType.olMailItem));
		original.Subject = "Email.ToolKit wrapped mail";
		string path = Path.Combine(context.DirectoryPath, "wrapped.msg");
		original.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);

		OutlookMail? opened = session.OpenMailItemFile(path);

		Assert.That(opened, Is.Not.Null);
		Assert.That(opened!.GetSynopses(), Does.Contain(original.Subject));
		original.Close(Outlook.OlInspectorClose.olDiscard);
	}

	/// <summary>
	/// Verifies a non-mail shared item is rejected by the mail wrapper.
	/// </summary>
	[Test]
	public void OpenMailItemFileRejectsAppointment()
	{
		using OutlookTestContext context = new();
		OutlookSession session = new(context.NameSpace);
		Outlook.AppointmentItem original = context.Track(
			(Outlook.AppointmentItem)context.Application.CreateItem(
				Outlook.OlItemType.olAppointmentItem));
		original.Subject = "Email.ToolKit appointment";
		string path = Path.Combine(context.DirectoryPath, "appointment.msg");
		original.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);

		Assert.That(session.OpenMailItemFile(path), Is.Null);
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
}
