/////////////////////////////////////////////////////////////////////////////
// <copyright file="EmailToolKitTests.cs" company="James John McGuire">
// Copyright © 2021 - 2024 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using DigitalZenWorks.Common.Utilities;
using Microsoft.Office.Interop.Outlook;
using NUnit.Framework;
using NUnit.Framework.Internal;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

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
		private DirectoryInfo testFolder;

		/// <summary>
		/// One time set up method.
		/// </summary>
		[OneTimeSetUp]
		public void OneTimeSetUp()
		{
			outlookAccount = OutlookAccount.Instance;

			pstOutlook = new (outlookAccount);

			testFolder = Directory.CreateDirectory("TestFolder");

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

			bool result = Directory.Exists(testFolder.FullName);

			if (true == result)
			{
				Directory.Delete(testFolder.FullName, true);
			}
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

			Assert.That(folder, Is.Not.Null);

			folder.Delete();
			Marshal.ReleaseComObject(folder);
		}

		/// <summary>
		/// Test for create pst store.
		/// </summary>
		[Test]
		public void TestCreatePstStore()
		{
			Assert.That(store, Is.Not.Null);
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
			Assert.That(exists, Is.False);

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
			Assert.That(exists, Is.True);

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
			mailItem = mailItem.Move(mainFolder);

			MailItem mailItem2 = outlookAccount.CreateMailItem(
				"someoneelse@example.com",
				"This is another subject",
				"This is the message.");
			mailItem2 = mailItem2.Move(mainFolder);

			mailItem.Save();
			mailItem2.Save();

			string tester = mailItem.EntryID;
			string tester2 = mailItem2.EntryID;

			Assert.That(tester2, Is.Not.EqualTo(tester));

			// Clean up
			mailItem.Delete();
			mailItem2.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mailItem2);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test EML to PST with sucess.
		/// </summary>
		[Test]
		public void TestEmlToPstSuccess()
		{
			MAPIFolder rootFolder = store.GetRootFolder();

			string path = Path.Combine(testFolder.FullName, "TestEmail.eml");
			bool result = FileUtils.CreateFileFromEmbeddedResource(
				"ToolKit.Tests.TestEmail.eml", path);

			Assert.That(result, Is.True);

			Migrate.EmlToPst(path, storePath, true);

			string baseName =
				Path.GetFileNameWithoutExtension(storePath);

			bool exists =
				OutlookFolder.DoesFolderExist(rootFolder, baseName);
			Assert.That(exists, Is.True);

			MAPIFolder folder =
				OutlookFolder.GetSubFolder(rootFolder, baseName);
			Assert.That(folder, Is.Not.Null);

			int count = folder.Items.Count;
			Assert.That(count, Is.GreaterThan(0));
		}

		/// <summary>
		/// Test to check if trying to get a folder with a name in a
		/// different case sensitivity fails.
		/// </summary>
		[Test]
		public void TestGetSubFolderCaseSensitiveFail()
		{
			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			// Create test sub folders
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(mainFolder, "Testing");
			Marshal.ReleaseComObject(subFolder);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "testing", true);

			Assert.That(subFolder, Is.Null);

			// Clean up
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test to check if trying to get a folder with a name in a
		/// different case sensitivity fails.
		/// </summary>
		[Test]
		public void TestGetSubFolderCaseSensitiveTrue()
		{
			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();
			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			// Create test sub folders
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(mainFolder, "Testing");
			Marshal.ReleaseComObject(subFolder);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Testing", true);

			Assert.That(subFolder, Is.Not.Null);

			// Clean up
			Marshal.ReleaseComObject(subFolder);
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

			Assert.That(afterHtmlBody, Is.EqualTo(htmlBody));
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

			Assert.That(htmlBody, Is.EqualTo(afterHtmlBody));
		}

		/// <summary>
		/// Test for comparing two different MailItems by content.
		/// </summary>
		[Test]
		public void TestHtmlBodyTrimLineEndingsNoChange()
		{
			string htmlBody = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
				"&nbsp;&nbsp;&nbsp;" +
				"<BR>\r\n</FONT>\r\n</P>\r\n</BODY>\r\n</HTML>";
			string afterHtmlBody = htmlBody;

			htmlBody = HtmlEmail.Trim(htmlBody);

			Assert.That(htmlBody, Is.EqualTo(afterHtmlBody));
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
			mailItem = mailItem.Move(mainFolder);

			MailItem mailItem2 = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is aka subject",
				"This is the message.");
			mailItem2 = mailItem2.Move(mainFolder);

			OutlookItem outlookItem = new (mailItem);
			string hash = outlookItem.Hash;

			OutlookItem outlookItem2 = new (mailItem2);
			string hash2 = outlookItem2.Hash;

			Assert.That(hash2, Is.Not.EqualTo(hash));

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

			string path = Path.Combine(testFolder.FullName, "TestEmail.eml");
			bool result = FileUtils.CreateFileFromEmbeddedResource(
				"ToolKit.Tests.TestEmail.eml", path);

			Assert.That(result, Is.True);

			MailItem mailItem = Migrate.EmlFileToPst(path, storePath);
			MailItem mailItem2 = Migrate.EmlFileToPst(path, storePath);

			OutlookItem outlookItem = new (mailItem);
			string hash = outlookItem.Hash;

			OutlookItem outlookItem2 = new (mailItem2);
			string hash2 = outlookItem2.Hash;

			Assert.That(hash2, Is.EqualTo(hash));

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
			mailItem = mailItem.Move(mainFolder);

			OutlookItem outlookItem = new (mailItem);
			string hash = outlookItem.Hash;

			OutlookItem outlookItem2 = new (mailItem);
			string hash2 = outlookItem2.Hash;

			Assert.That(hash2, Is.EqualTo(hash));

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

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(rootFolder, false);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Testing (1)");

			Assert.That(subFolder, Is.Null);

			// Clean up
			mailItem.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folders.
		/// </summary>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// unit test.</returns>
		[Test]
		public async Task TestMergeFoldersAsync()
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

			OutlookFolder outlookFolder = new (outlookAccount);
			await outlookFolder.MergeFoldersAsync(rootFolder, false).
				ConfigureAwait(false);

			await Task.Delay(200).ConfigureAwait(false);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Testing (1)");

			Assert.That(subFolder, Is.Null);

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

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(rootFolder, false);

			System.Threading.Thread.Sleep(200);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Testing_5");
			Assert.That(subFolder, Is.Null);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "_Testing");
			Assert.That(subFolder, Is.Null);

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
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// unit test.</returns>
		[Test]
		public async Task TestMergeFoldersAggresiveAsync()
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

			OutlookFolder outlookFolder = new (outlookAccount);
			await outlookFolder.MergeFoldersAsync(rootFolder, false).
				ConfigureAwait(false);

			await Task.Delay(200).ConfigureAwait(false);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Testing_5");
			Assert.That(subFolder, Is.Null);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "_Testing");
			Assert.That(subFolder, Is.Null);

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
				"2023",
				"This is the subject");

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(rootFolder, false);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "2023");

			Assert.That(subFolder, Is.Not.Null);

			// Clean up
			mailItem.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folders.
		/// </summary>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// unit test.</returns>
		[Test]
		public async Task TestMergeFoldersAllNumbersFolderAsync()
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
				"2023",
				"This is the subject");

			OutlookFolder outlookFolder = new (outlookAccount);
			await outlookFolder.MergeFoldersAsync(rootFolder, false).
				ConfigureAwait(false);

			await Task.Delay(200).ConfigureAwait(false);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "2023");

			Assert.That(subFolder, Is.Not.Null);

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

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(rootFolder, false);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Main Test Folder");

			Assert.That(subFolder, Is.Null);

			// Clean up
			mailItem.Delete();
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folders.
		/// </summary>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// unit test.</returns>
		[Test]
		public async Task TestMergeFolderWithParentAsync()
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

			OutlookFolder outlookFolder = new (outlookAccount);
			await outlookFolder.MergeFoldersAsync(rootFolder, false).
				ConfigureAwait(false);

			await Task.Delay(200).ConfigureAwait(false);

			subFolder =
				OutlookFolder.GetSubFolder(mainFolder, "Main Test Folder");

			Assert.That(subFolder, Is.Null);

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

			string path = Path.Combine(testFolder.FullName, "TestEmail.eml");
			bool result = FileUtils.CreateFileFromEmbeddedResource(
				"ToolKit.Tests.TestEmail.eml", path);

			Assert.That(result, Is.True);

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");

			MailItem mailItem2 = Migrate.EmlFileToPst(path, storePath);
			MailItem mailItem3 = Migrate.EmlFileToPst(path, storePath);

			mailItem = mailItem.Move(mainFolder);
			mailItem2 = mailItem2.Move(mainFolder);
			mailItem3 = mailItem3.Move(mainFolder);

			OutlookFolder outlookFolder = new (outlookAccount);
			int removedDuplicates =
				outlookFolder.RemoveDuplicates(mainFolder, false);

			Assert.That(removedDuplicates, Is.EqualTo(1));

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

			Assert.That(subFolder, Is.Null);

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

			OutlookFolder.RemoveEmptyFolders(rootFolder, true);

			subFolder = OutlookFolder.GetSubFolder(
				rootFolder, "Temporary Test Folder");

			Assert.That(subFolder, Is.Null);

			if (subFolder != null)
			{
				Marshal.ReleaseComObject(subFolder);
			}

			Marshal.ReleaseComObject(rootFolder);
		}

		/// <summary>
		/// Test for removing empty folder.
		/// </summary>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// unit test.</returns>
		[Test]
		public async Task TestRemoveEmptyFoldersAsync()
		{
			MAPIFolder rootFolder = store.GetRootFolder();

			MAPIFolder subFolder = OutlookFolder.AddFolder(
				rootFolder, "Temporary Test Folder");
			Marshal.ReleaseComObject(subFolder);

			await OutlookFolder.RemoveEmptyFoldersAsync(rootFolder, true).
				ConfigureAwait(false);

			subFolder = OutlookFolder.GetSubFolder(
				rootFolder, "Temporary Test Folder");

			Assert.That(subFolder, Is.Null);

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

			header = OutlookItem.RemoveMimeOleVersion(header);

			Assert.That(header, Is.EqualTo(afterHeader));
		}

		/// <summary>
		/// Test for checking if RtfEmail.Trim is correct.
		/// </summary>
		[Test]
		public void TestRftBodyTrim()
		{
			byte[] sampleBytes =
			[
				32, 32, 32, 32, 32, 32, 92, 112, 97, 114, 13, 10, 92, 112, 97,
				114, 13, 10, 92, 112, 97, 114, 13, 10, 92, 112, 97, 114, 13,
				10, 92, 112, 97, 114, 13, 10, 92, 112, 97, 114, 13, 10, 92,
				112, 97, 114, 13, 10, 92, 112, 97, 114, 13, 10, 125, 13, 10, 0
			];
			byte[] afterBytes =
			[
				32, 32, 32, 32, 32, 32, 92, 112, 97, 114, 13, 10, 125, 13,
				10, 0
			];

			sampleBytes = RtfEmail.Trim(sampleBytes);

			Assert.That(afterBytes, Is.EqualTo(sampleBytes));
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

			mailItem = mailItem.Move(subFolder);

			Marshal.ReleaseComObject(subFolder);

			return mailItem;
		}
	}
}
