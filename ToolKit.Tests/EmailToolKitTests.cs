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
		private OutlookAccount outlookAccount;
		private OutlookStore pstOutlook;
		private Store store;
		private string storePath;

		/// <summary>
		/// One time set up method.
		/// </summary>
		[OneTimeSetUp]
		public void OneTimeSetUp()
		{
			outlookAccount = OutlookAccount.Instance;

			pstOutlook = new (outlookAccount);

			string fileName = Path.GetTempFileName();

			// A 0 byte sized file is created.  Need to remove it.
			File.Delete(fileName);
			storePath = Path.ChangeExtension(fileName, ".pst");

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

			store = outlookAccount.GetStore(storePath);
		}

		/// <summary>
		/// One time tear down method.
		/// </summary>
		[OneTimeTearDown]
		public void OneTimeTearDown()
		{
			OutlookStore.EmptyDeletedItemsFolder(store);
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
		/// Test for creating folder from path.
		/// </summary>
		[Test]
		public void TestCreateFolderPath()
		{
			MAPIFolder folder =
				OutlookFolder.CreateFolderPath(store, "Testing/Test");

			Assert.NotNull(folder);

			folder.Delete();
			Marshal.ReleaseComObject(folder);
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
		/// Test for does folder exist.
		/// </summary>
		[Test]
		public void TestDoesFolderExistFalse()
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			bool exists =
				OutlookFolder.DoesFolderExist(mainFolder, "Some Sub Folder");
			Assert.False(exists);

			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for does folder exist.
		/// </summary>
		[Test]
		public void TestDoesFolderExistTrue()
		{
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");
			MAPIFolder subFolder = OutlookFolder.AddFolder(
				mainFolder, "Some Sub Folder");

			bool exists =
				OutlookFolder.DoesFolderExist(mainFolder, "Some Sub Folder");
			Assert.True(exists);

			Marshal.ReleaseComObject(subFolder);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
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

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = outlookAccount.CreateMailItem(
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
		public void TestHtmlBodyTrimBr()
		{
			string htmlBody = "<BR>\r\n<BR>\r\n<BR>\r\n<BR>\r\n</FONT>\r\n" +
				"</P>\r\n\r\n</BODY>\r\n</HTML>";
			string afterHtmlBody =
				"<BR>\r\n</FONT>\r\n</P>\r\n</BODY>\r\n</HTML>";

			htmlBody = HtmlEmail.Trim(htmlBody);

			Assert.AreEqual(htmlBody, afterHtmlBody);
		}

		/// <summary>
		/// Test for comparing two different MailItems by content.
		/// </summary>
		[Test]
		public void TestHtmlBodyTrimLineEndings()
		{
			string htmlBody = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
				"&nbsp;&nbsp;&nbsp;<BR>\r\n<BR>\r\n<BR>\r\n<BR>\r\n<BR>\r\n" +
				"<BR>\r\n<BR>\r\n</FONT>\r\n</P>\r\n\r\n</BODY>\r\n</HTML>";

			string afterHtmlBody = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
				"&nbsp;&nbsp;&nbsp;" +
				"<BR>\r\n</FONT>\r\n</P>\r\n</BODY>\r\n</HTML>";

			htmlBody = HtmlEmail.Trim(htmlBody);

			Assert.AreEqual(afterHtmlBody, htmlBody);
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

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = outlookAccount.CreateMailItem(
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

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = outlookAccount.CreateMailItem(
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

			MailItem mailItem = outlookAccount.CreateMailItem(
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

			MailItem mailItem = AddFolderAndMessage(
				outlookAccount,
				mainFolder,
				"Testing (1)",
				"This is the subject");

			// Review
			storePath = OutlookStore.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(path, rootFolder, false);

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
		/// Test for removing empty folders.
		/// </summary>
		[Test]
		public void TestMergeFoldersAggresive()
		{
			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			// Create sub folders
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(mainFolder, "Testing");
			Marshal.ReleaseComObject(subFolder);

			MailItem mailItem = AddFolderAndMessage(
				outlookAccount,
				mainFolder,
				"Testing_5",
				"This is the subject 1");

			MailItem mailItem2 = AddFolderAndMessage(
				outlookAccount,
				mainFolder,
				"_Testing",
				"This is the subject 3");

			// Review
			storePath = OutlookStore.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(path, rootFolder, false);

			System.Threading.Thread.Sleep(200);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Testing_5");
			Assert.IsNull(subFolder);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "_Testing");
			Assert.IsNull(subFolder);

			// Clean up
			mailItem.Delete();
			mailItem2.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mailItem2);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folders.
		/// </summary>
		[Test]
		public void TestMergeFoldersAllNumbersFolder()
		{
			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			// Create sub folders
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(mainFolder, "Testing");
			Marshal.ReleaseComObject(subFolder);

			MailItem mailItem = AddFolderAndMessage(
				outlookAccount,
				mainFolder,
				"2022",
				"This is the subject");

			// Review
			storePath = OutlookStore.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(path, rootFolder, false);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "2022");

			Assert.IsNotNull(subFolder);

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
		public void TestMergeFolderWithParent()
		{
			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			// Create sub folders
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(mainFolder, "Main Test Folder");
			Marshal.ReleaseComObject(subFolder);

			MailItem mailItem = AddFolderAndMessage(
				outlookAccount,
				mainFolder,
				"Main Test Folder",
				"This is the subject");

			// Review
			storePath = OutlookStore.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(path, rootFolder, false);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Main Test Folder");

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

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(mainFolder);

			MailItem mailItem2 = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is aka subject",
				"This is the message.");
			mailItem2.Move(mainFolder);

			MailItem mailItem3 = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem3.Move(mainFolder);

			string storePath = OutlookStore.GetStoreName(store);

			OutlookFolder outlookFolder = new (outlookAccount);
			int[] counts =
				outlookFolder.RemoveDuplicates(storePath, mainFolder, false);

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

			OutlookStore.RemoveFolder(rootFolder.Name, subFolder, false);

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

			storePath = OutlookStore.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			OutlookFolder.RemoveEmptyFolders(path, rootFolder, true);

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
		/// Test for comparing two different MailItems by content.
		/// </summary>
		[Test]
		public void TestRemoveMimeOleVersion()
		{
			string header = @"X-Priority: 3\r\nX-MSMail-Priority: Normal\r\n" +
				@"X-MimeOLE: Produced By Microsoft MimeOLE V6.00.2900.2180" +
				@"\r\nFrom: <admin@example.com>\r\nTo: " +
				@"<somebody@example.com>,\t<somebodyelse@example.com>\r\n" +
				@"Subject: Subject Statement\r\n";
			string afterHeader = @"X-Priority: 3\r\nX-MSMail-Priority: " +
				@"Normal\r\nX-MimeOLE: Produced By Microsoft MimeOLE" +
				@"\r\nFrom: <admin@example.com>\r\nTo: " +
				@"<somebody@example.com>,\t<somebodyelse@example.com>\r\n" +
				@"Subject: Subject Statement\r\n";

			header = MapiItem.RemoveMimeOleVersion(header);

			Assert.AreEqual(afterHeader, header);
		}

		/// <summary>
		/// Test for checking if RtfEmail.Trim is correct.
		/// </summary>
		[Test]
		public void TestRftBodyTrim()
		{
			byte[] sampleBytes = new byte[]
			{
				32, 32, 32, 32, 32, 32, 92, 112, 97, 114, 13, 10, 92, 112, 97,
				114, 13, 10, 92, 112, 97, 114, 13, 10, 92, 112, 97, 114, 13,
				10, 92, 112, 97, 114, 13, 10, 92, 112, 97, 114, 13, 10, 92,
				112, 97, 114, 13, 10, 92, 112, 97, 114, 13, 10, 125, 13, 10, 0
			};
			byte[] afterBytes = new byte[]
			{
				32, 32, 32, 32, 32, 32, 92, 112, 97, 114, 13, 10, 125, 13,
				10, 0
			};

			sampleBytes = RtfEmail.Trim(sampleBytes);

			Assert.AreEqual(sampleBytes, afterBytes);
		}

		/// <summary>
		/// Test for sanity check.
		/// </summary>
		[Test]
		public void TestSanityCheck()
		{
			Assert.Pass();
		}

		private static MailItem AddFolderAndMessage(
			OutlookAccount outlookAccount,
			MAPIFolder parentFolder,
			string folderName,
			string subject)
		{
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(parentFolder, folderName);

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				subject,
				"This is the message.");
			mailItem.UnRead = false;
			mailItem.Save();

			mailItem.Move(subFolder);

			Marshal.ReleaseComObject(subFolder);

			return mailItem;
		}
	}
}
