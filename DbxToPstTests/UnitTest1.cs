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
		public void TestIndexedInfo()
		{
			byte[] testBytes =
			{
				0x00, 0x00, 0x00, 0x00, 0x38, 0x00, 0x00, 0x00, 0x00, 0x00,
				0x04, 0x01, 0x80, 0x11, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00,
				0x05, 0x1d, 0x00, 0x00, 0x86, 0x29, 0x00, 0x04, 0x64, 0x69,
				0x73, 0x63, 0x75, 0x73, 0x73, 0x69, 0x6f, 0x6e, 0x2e, 0x66,
				0x61, 0x73, 0x74, 0x61, 0x6E, 0x64, 0x66, 0x75, 0x72, 0x69,
				0x75, 0x73, 0x2e, 0x63, 0x6f, 0x6d, 0x00, 0x30, 0x30, 0x30,
				0x30, 0x30, 0x30, 0x31, 0x37, 0x00, 0x00, 0x00
			};

			DbxIndexedItem item = new ();
			item.ReadIndex(testBytes, 0);

			uint value = item.GetValue(DbxFolderIndexedItem.Id);
			Assert.AreEqual(value, 0x11);

			value = item.GetValue(DbxFolderIndexedItem.ParentId);
			Assert.AreEqual(value, 0);

			string name = item.GetString(DbxFolderIndexedItem.Name);
			string expected = "discussion.fastandfurius.com";
			Assert.That(name, Is.EqualTo(expected));
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