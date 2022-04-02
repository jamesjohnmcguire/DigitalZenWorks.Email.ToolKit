/////////////////////////////////////////////////////////////////////////////
// <copyright file="EmailToolKitTests.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using DigitalZenWorks.Email.ToolKit;
using Microsoft.Office.Interop.Outlook;
using NUnit.Framework;
using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

[assembly: CLSCompliant(true)]

namespace DigitalZenWorks.Email.ToolKit.Tests
{
	/// <summary>
	/// Test class.
	/// </summary>
	public class EmailToolKitTests
	{
		private OutlookStorage pstOutlook;
		private Store store;
		private string storePath;

		/// <summary>
		/// One time set up method.
		/// </summary>
		[OneTimeSetUp]
		public void OneTimeSetUp()
		{
			string fileName = Path.GetTempFileName();

			// A 0 byte sized file is created.  Need to remove it.
			File.Delete(fileName);
			storePath = Path.ChangeExtension(fileName, ".pst");

			pstOutlook = new ();

			// PST provider in Outlook keeps the PST file open for 30 minutes
			// after closing it for the performance reasons. So, try to delete
			// it now, as it may be more than 30 minutes since last access.
			bool exists = File.Exists(storePath);

			if (exists == true)
			{
				try
				{
					File.Delete(storePath);
				}
				catch (IOException)
				{
				}
			}

			store = pstOutlook.GetStore(storePath);
		}

		/// <summary>
		/// One time tear down method.
		/// </summary>
		[OneTimeTearDown]
		public void OneTimeTearDown()
		{
			pstOutlook.EmptyDeletedItemsFolder();
			pstOutlook.RemoveStore(store);
		}

		/// <summary>
		/// Set up method.
		/// </summary>
		[SetUp]
		public void Setup()
		{
		}

		/// <summary>
		/// Test for create pst store.
		/// </summary>
		[Test]
		public void TestCreatePstStore()
		{
			Assert.NotNull(store);
		}

		/// <summary>
		/// Test for checking of duplicate items.
		/// </summary>
		[Test]
		public void TestDifferentItemsEntryIds()
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			MailItem mailItem = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = pstOutlook.CreateMailItem(
				"someoneelse@example.com",
				"This is another subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			mailItem.Save();
			mailItem2.Save();

			string tester = mailItem.EntryID;
			string tester2 = mailItem2.EntryID;

			Assert.AreNotEqual(tester, tester2);

			// Clean up
			mailItem.Delete();
			mailItem2.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mailItem2);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for comparing two different MailItems by content.
		/// </summary>
		[Test]
		public void TestMailItemsAreNotSameByContent()
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			MailItem mailItem = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is aka subject",
				"This is the message.");
			mailItem2.Move(mainFolder);

			string path = OutlookFolder.GetFolderPath(mainFolder);
			string hash = MapiItem.GetItemHash(path, mailItem);
			string hash2 = MapiItem.GetItemHash(path, mailItem2);

			Assert.AreNotEqual(hash, hash2);

			// Clean up
			mailItem.Delete();
			mailItem2.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mailItem2);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for comparing two MailItems by content.
		/// </summary>
		[Test]
		public void TestMailItemsSameByContent()
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			MailItem mailItem = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem2.Move(mainFolder);

			string path = OutlookFolder.GetFolderPath(mainFolder);
			string hash = MapiItem.GetItemHash(path, mailItem);
			string hash2 = MapiItem.GetItemHash(path, mailItem2);

			Assert.AreEqual(hash, hash2);

			// Clean up
			mailItem.Delete();
			mailItem2.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mailItem2);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for comparing two MailItems by refence.
		/// </summary>
		/// <remarks>This is more of a sanity check.</remarks>
		[Test]
		public void TestMailItemsSameByReference()
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			MailItem mailItem = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			string path = OutlookFolder.GetFolderPath(mainFolder);
			string hash = MapiItem.GetItemHash(path, mailItem);
			string hash2 = MapiItem.GetItemHash(path, mailItem);

			Assert.AreEqual(hash, hash2);

			// Clean up
			mailItem.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folders.
		/// </summary>
		[Test]
		public void TestMergeFolders()
		{
			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			// Create sub folders
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(mainFolder, "Testing");
			Marshal.ReleaseComObject(subFolder);

			subFolder = OutlookFolder.AddFolder(mainFolder, "Testing (1)");

			MailItem mailItem = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(subFolder);

			Marshal.ReleaseComObject(subFolder);

			// Review
			storePath = OutlookStorage.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			OutlookFolder outlookFolder = new ();
			outlookFolder.MergeFolders(path, rootFolder);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Testing (1)");

			Assert.IsNull(subFolder);

			// Clean up
			mailItem.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folder.
		/// </summary>
		[Test]
		public void TestRemoveDuplicates()
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Duplicates Test Folder");

			MailItem mailItem = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is aka subject",
				"This is the message.");
			mailItem2.Move(mainFolder);

			MailItem mailItem3 = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem3.Move(mainFolder);

			int[] counts =
				pstOutlook.RemoveDuplicates(mainFolder, false, true);

			Assert.AreEqual(counts[0], 1);
			Assert.AreEqual(counts[1], 2);

			// Clean up
			mailItem.Delete();
			mailItem2.Delete();
			mailItem3.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folder.
		/// </summary>
		[Test]
		public void TestRemoveEmptyFolder()
		{
			MAPIFolder rootFolder = store.GetRootFolder();

			MAPIFolder subFolder = OutlookFolder.AddFolder(
				rootFolder, "Temporary Test Folder");

			pstOutlook.RemoveFolder(rootFolder.Name, subFolder, false);

			Marshal.ReleaseComObject(subFolder);

			System.Threading.Thread.Sleep(200);
			subFolder = OutlookFolder.GetSubFolder(
				rootFolder, "Temporary Test Folder");

			Assert.IsNull(subFolder);

			if (subFolder != null)
			{
				Marshal.ReleaseComObject(subFolder);
			}

			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folder.
		/// </summary>
		[Test]
		public void TestRemoveEmptyFolders()
		{
			MAPIFolder rootFolder = store.GetRootFolder();

			MAPIFolder subFolder = OutlookFolder.AddFolder(
				rootFolder, "Temporary Test Folder");
			Marshal.ReleaseComObject(subFolder);

			storePath = OutlookStorage.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			pstOutlook.RemoveEmptyFolders(path, rootFolder);

			subFolder = OutlookFolder.GetSubFolder(
				rootFolder, "Temporary Test Folder");

			Assert.IsNull(subFolder);

			if (subFolder != null)
			{
				Marshal.ReleaseComObject(subFolder);
			}

			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for sanity check.
		/// </summary>
		[Test]
		public void TestSanityCheck()
		{
			Assert.Pass();
		}
	}
}
