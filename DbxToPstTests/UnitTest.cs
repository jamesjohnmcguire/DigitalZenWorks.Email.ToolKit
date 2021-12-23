using DbxToPstLibrary;
using Microsoft.Office.Interop.Outlook;
using NUnit.Framework;
using System;
using System.IO;

[assembly: CLSCompliant(true)]

namespace DigitalZenWorks.Email.DbxToPstTests.Tests
{
	/// <summary>
	/// Test class.
	/// </summary>
	public class DbxToPstTestsTests
	{
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
			string basePath = Path.GetTempPath();
			string path = basePath + "Test.pst";
			PstOutlook pstOutlook = new ();

			// PST provider in Outlook keeps the PST file open for 30 minutes
			// after closing it for the performance reasons.
			try
			{
				File.Delete(path);

			}
			catch (IOException) { }

			Store store = pstOutlook.CreateStore(path);

			Assert.NotNull(store);

			pstOutlook.RemoveStore(store);
		}
	}
}
