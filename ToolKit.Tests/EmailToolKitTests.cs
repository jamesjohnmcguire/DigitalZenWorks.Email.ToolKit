/////////////////////////////////////////////////////////////////////////////
// <copyright file="EmailToolKitTests.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using DigitalZenWorks.Email.ToolKit;
using Microsoft.Office.Interop.Outlook;
using NUnit.Framework;
using System;
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
			string basePath = Path.GetTempPath();
			storePath = basePath + "Test.pst";

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

			store = pstOutlook.CreateStore(storePath);
		}

		/// <summary>
		/// One time tear down method.
		/// </summary>
		[OneTimeTearDown]
		public void OneTimeTearDown()
		{
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
		/// Test for sanity check.
		/// </summary>
		[Test]
		public void TestSanityCheck()
		{
			Assert.Pass();
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
		/// Test for removing empty folder.
		/// </summary>
		[Test]
		public void TestRemoveEmptyFolder()
		{
			MAPIFolder rootFolder = store.GetRootFolder();

			MAPIFolder subFolder = OutlookStorage.AddFolderSafe(
				rootFolder, "Temporary Test Folder");

			pstOutlook.RemoveFolder(rootFolder.Name, subFolder, false);

			Marshal.ReleaseComObject(subFolder);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookStorage.GetSubFolder(rootFolder, "Temporary Test Folder");

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

			MAPIFolder subFolder = OutlookStorage.AddFolderSafe(
				rootFolder, "Temporary Test Folder");
			Marshal.ReleaseComObject(subFolder);

			storePath = OutlookStorage.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			pstOutlook.RemoveEmptyFolders(path, rootFolder);

			subFolder =
				OutlookStorage.GetSubFolder(rootFolder, "Temporary Test Folder");

			Assert.IsNull(subFolder);

			if (subFolder != null)
			{
				Marshal.ReleaseComObject(subFolder);
			}

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
			MAPIFolder mainFolder = OutlookStorage.AddFolderSafe(
				rootFolder, "Main Test Folder");

			// Create sub folders
			MAPIFolder subFolder =
				OutlookStorage.AddFolderSafe(mainFolder, "Testing");
			Marshal.ReleaseComObject(subFolder);

			subFolder = OutlookStorage.AddFolderSafe(mainFolder, "Testing (1)");

			MailItem mailItem = pstOutlook.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem.Move(subFolder);

			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(subFolder);

			// Review
			storePath = OutlookStorage.GetStoreName(store) + "::";
			string path = storePath + rootFolder.Name;

			pstOutlook.MergeFolders(path, rootFolder);

			System.Threading.Thread.Sleep(200);
			subFolder =
				OutlookStorage.GetSubFolder(mainFolder, "Testing (1)");

			Assert.IsNull(subFolder);

			// Clean up
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}
	}
}
