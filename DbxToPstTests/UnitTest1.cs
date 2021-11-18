using NUnit.Framework;
using DbxToPstLibrary;
using System;

namespace DbxToPstTests
{
	public class Tests
	{
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
		public void TestDbxToPst()
		{
			string applicationDataDirectory = @"DigitalZenWorks\DbxToPst";
			string baseDataDirectory = Environment.GetFolderPath(
				Environment.SpecialFolder.ApplicationData,
				Environment.SpecialFolderOption.Create) + @"\" +
				applicationDataDirectory;

			string path = baseDataDirectory + "\\TestFolder";
			bool result = Migrate.DbxToPst(path);
			Assert.IsTrue(result);
		}

	}
}