using DbxToPstLibrary;
using Microsoft.Office.Interop.Outlook;
using NUnit.Framework;
using System;

namespace DbxToPstTests
{
	public class Tests
	{
		private static readonly string applicationDataDirectory =
			@"DigitalZenWorks\DbxToPst";
		private static readonly string baseDataDirectory =
			Environment.GetFolderPath(
				Environment.SpecialFolder.ApplicationData,
				Environment.SpecialFolderOption.Create) + @"\" +
				applicationDataDirectory;

		[SetUp]
		public void Setup()
		{
		}

		[Test]
		public void TestSanityCheck()
		{
			Assert.Pass();
		}

		[Test]
		public void TestCreatePstStore()
		{
			string path = baseDataDirectory + "\\NewTest.pst";
			PstOutlook pstOutlook = new ();

			Microsoft.Office.Interop.Outlook.Store store = pstOutlook.CreateStore(path);

			Assert.NotNull(store);
		}

		[Test]
		public void TestDbxToPst()
		{
			string path = baseDataDirectory + "\\TestFolder";
			bool result = Migrate.DbxToPst(path);
			Assert.IsTrue(result);
		}

	}
}