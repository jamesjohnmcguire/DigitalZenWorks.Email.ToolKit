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
		private PstOutlook pstOutlook;
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
			try
			{
				File.Delete(storePath);
			}
			catch (IOException)
			{
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
	}
}
