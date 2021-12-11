using DbxToPstLibrary;
using Microsoft.Office.Interop.Outlook;
using NUnit.Framework;
using System;

[assembly: CLSCompliant(true)]

namespace DigitalZenWorks.Email.DbxToPstTests.Tests
{
	/// <summary>
	/// Test class.
	/// </summary>
	public class DbxToPstTestsTests
	{
		private const string applicationDataDirectory =
			@"DigitalZenWorks\DbxToPst";
		private static readonly string baseDataDirectory =
			Environment.GetFolderPath(
				Environment.SpecialFolder.ApplicationData,
				Environment.SpecialFolderOption.Create) + @"\" +
				applicationDataDirectory;

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
			string path = baseDataDirectory + "\\NewTest.pst";
			PstOutlook pstOutlook = new ();

			Microsoft.Office.Interop.Outlook.Store store = pstOutlook.CreateStore(path);

			Assert.NotNull(store);

			pstOutlook.RemoveStore(store);
		}
	}
}
