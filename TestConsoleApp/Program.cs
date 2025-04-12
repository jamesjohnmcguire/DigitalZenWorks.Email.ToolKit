﻿/////////////////////////////////////////////////////////////////////////////
// <copyright file="Program.cs" company="James John McGuire">
// Copyright © 2021 - 2025 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Email.DbxOutlookExpress;
using DigitalZenWorks.Email.ToolKit;
using Microsoft.Office.Interop.Outlook;
using MsgKit;
using Serilog;
using Serilog.Configuration;
using Serilog.Events;
using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

using CommonLogging = Common.Logging;

[assembly: CLSCompliant(true)]
#pragma warning disable IDE0051
#pragma warning disable IDE0059

namespace DigitalZenWorks.Email.ToolKit.Test
{
	/// <summary>
	/// Dbx to pst test program class.
	/// </summary>
	internal static class Program
	{
		private const string ApplicationDataDirectory =
			@"DigitalZenWorks\DbxToPst";

		private static readonly string BaseDataDirectory =
			Environment.GetFolderPath(
				Environment.SpecialFolder.ApplicationData,
				Environment.SpecialFolderOption.Create) + @"\" +
				ApplicationDataDirectory;

		private static readonly ILog Log = LogManager.GetLogger(
#pragma warning disable CS8602 // Dereference of a possibly null reference.
			MethodBase.GetCurrentMethod().DeclaringType);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

		/// <summary>
		/// The program's main entry point.
		/// </summary>
		/// <param name="arguments">The arguments given to the program.</param>
		public static void Main(string[] arguments)
		{
			LogInitialization();

			Log.Info("Test console app");

			TestRegex();

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			TestTargetFrameworks();

			TestDeleteCalendarItem(outlookAccount);
			TestGetHash(outlookAccount);

			TestMsgCompare(outlookAccount);

			TestMergeFolders(outlookAccount);

			Encoding.RegisterProvider(
				CodePagesEncodingProvider.Instance);
			Encoding encoding = Encoding.GetEncoding("shift_jis");

			string path = BaseDataDirectory + @"\TestFolder";

			TestConvertToMsgFile(path, encoding);

			TestTree();

			TestSetTree(path, encoding);
			TestListSet(path, encoding);

			if (arguments != null && arguments.Length > 0)
			{
			}
			else
			{
				Log.Error("Invalid arguments");
			}
		}

		private static void LogInitialization()
		{
			string logFilePath = BaseDataDirectory + "\\DbxToPst.log";
			string outputTemplate =
				"[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] " +
				"{Message:lj}{NewLine}{Exception}";

			LoggerConfiguration configuration = new ();
			LoggerSinkConfiguration sinkConfiguration = configuration.WriteTo;
			sinkConfiguration.Console(
				LogEventLevel.Verbose,
				outputTemplate,
				CultureInfo.InvariantCulture);
			sinkConfiguration.File(
				logFilePath,
				LogEventLevel.Verbose,
				outputTemplate,
				CultureInfo.InvariantCulture);
			Serilog.Log.Logger = configuration.CreateLogger();

			LogManager.Adapter =
				new CommonLogging.Serilog.SerilogFactoryAdapter();
		}

		private static void TestConvertToMsgFile(
			string path, Encoding encoding)
		{
			DbxMessagesFile messagesFile = new (path, encoding);

			DbxMessage message = messagesFile.GetMessage(79);

			Stream dbxStream = message.MessageStream;

			string msgPath = BaseDataDirectory + @"\test.msg";

			File.Delete(msgPath);

			using Stream msgStream =
				new FileStream(msgPath, FileMode.Create);

			Converter.ConvertEmlToMsg(dbxStream, msgStream);
		}

		private static void TestDeleteCalendarItem(OutlookAccount outlookAccount)
		{
			string storePath = @"C:\Users\JamesMc\Data\ProgramData\Outlook\" +
				"Test.pst";

			Store store = outlookAccount.GetStore(storePath);

			MAPIFolder rootFolder = store.GetRootFolder();

			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Calendar");

			Items items = mainFolder.Items;

			// Office uses 1 based indexes from VBA.
			// Iterate in reverse order as the group may change.
			for (int index = items.Count; index > 0; index--)
			{
				object item = items[index];

				if (item != null)
				{
					AppointmentItem? appointmentItem = item as AppointmentItem;

					appointmentItem!.Delete();
				}
			}
		}

		private static void TestFolder(string path, Encoding encoding)
		{
			DbxFolder dbxFolder = new (path, "TmpHold", encoding);
		}

		private static void TestGetHash(OutlookAccount outlookAccount)
		{
			// Create test store.
			string basePath = Path.GetTempPath();
			string storePath = basePath + "Test.pst";

			Store store = outlookAccount.GetStore(storePath);

			// Create top level folders
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
			string hash2 = outlookItem.Hash;

			if (hash.Equals(hash2, StringComparison.Ordinal))
			{
				Log.Info("Hashes are the same");
			}
			else
			{
				Log.Info("Hashes are NOT the same");
			}

			MailItem mailItem2 = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem2 = mailItem2.Move(mainFolder);

			OutlookItem outlookItem2 = new (mailItem2);
			hash2 = outlookItem2.Hash;

			if (hash.Equals(hash2, StringComparison.Ordinal))
			{
				Log.Info("Hashes are the same");
			}
			else
			{
				Log.Info("Hashes are NOT the same");
			}

			MailItem mailItem3 = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is aka subject",
				"This is the message.");
			mailItem3 = mailItem3.Move(mainFolder);

			OutlookItem outlookItem3 = new (mailItem3);
			hash2 = outlookItem3.Hash;

			if (hash.Equals(hash2, StringComparison.Ordinal))
			{
				Log.Info("Hashes are the same");
			}
			else
			{
				Log.Info("Hashes are NOT the same");
			}

			// Clean up
			Marshal.ReleaseComObject(mailItem);
			Marshal.ReleaseComObject(mailItem2);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		private static void TestListMessagesFile(
			string path, Encoding encoding)
		{
			DbxMessagesFile messagesFile = new (path, encoding);

			messagesFile.List();
		}

		private static void TestListSet(string path, Encoding encoding)
		{
			DbxSet set = new (path, encoding);

			set.List();
		}

		private static void TestMergeFolders(OutlookAccount outlookAccount)
		{
			// Create test store.
			string basePath = Path.GetTempPath();
			string storePath = basePath + "Test.pst";

			Store store = outlookAccount.GetStore(storePath);

			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();

			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			// Create sub folders
			MAPIFolder subFolder =
				OutlookFolder.AddFolder(mainFolder, "Testing");
			OutlookFolder.AddFolder(subFolder, "Testing2");
			OutlookFolder.AddFolder(subFolder, "Testing2 (1)");

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem = mailItem.Move(subFolder);

			subFolder = OutlookFolder.AddFolder(
				mainFolder, "Testing (1)");
			OutlookFolder.AddFolder(subFolder, "Testing2");
			OutlookFolder.AddFolder(subFolder, "Testing2 (1)");

			OutlookFolder outlookFolder = new (outlookAccount);
			outlookFolder.MergeFolders(rootFolder, true);

			// Clean up
			Marshal.ReleaseComObject(subFolder);
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		private static void TestMsgCompare(OutlookAccount outlookAccount)
		{
			// Create test store.
			string basePath = Path.GetTempPath();
			string storePath = basePath + "Test.pst";

			Store store = outlookAccount.GetStore(storePath);

			// Create top level folders
			MAPIFolder rootFolder = store.GetRootFolder();

			MAPIFolder mainFolder = OutlookFolder.AddFolder(
				rootFolder, "Main Test Folder");

			MailItem mailItem = outlookAccount.CreateMailItem(
				"someone@example.com",
				"This is the subject",
				"This is the message.");
			mailItem = mailItem.Move(mainFolder);

			string msgPath = basePath + "test.msg";
			mailItem.SaveAs(msgPath);
			byte[] msg1 = File.ReadAllBytes(msgPath);

			msgPath = basePath + "test2.msg";
			mailItem.SaveAs(msgPath);
			byte[] msg2 = File.ReadAllBytes(msgPath);

			if (msg1.Equals(msg2))
			{
				Log.Info("Messages are the same");
			}
			else
			{
				Log.Info("Messages are NOT the same");
			}

			// Clean up
			Marshal.ReleaseComObject(mainFolder);
			Marshal.ReleaseComObject(rootFolder);
		}

		private static void TestRegex()
		{
			string input = "123ABC79";
			string pattern = @"\d+$";
			string result = Regex.Replace(
				input,
				pattern,
				string.Empty,
				RegexOptions.ExplicitCapture);

			input = "2000";
			pattern = @"\d+$";
			result = Regex.Replace(
				input,
				pattern,
				string.Empty,
				RegexOptions.ExplicitCapture);

			input = "123ABC7";
			pattern = @"(?<=\D+)\d+$";
			result = Regex.Replace(
				input,
				pattern,
				string.Empty,
				RegexOptions.ExplicitCapture);
		}

		private static void TestSetTree(string path, Encoding encoding)
		{
			path += @"\Folders.dbx";
			DbxFoldersFile foldersFile = new (path, encoding);

			foldersFile.SetTreeOrdered();
			foldersFile.List();
		}

		private static void TestStringToStream()
		{
			string test = "Testing 1-2-3";

			byte[] byteArray = Encoding.UTF8.GetBytes(test);
			MemoryStream stream = new (byteArray);

			TestStream(stream);
		}

		private static void TestStream(Stream stream)
		{
			using StreamReader reader = new (stream);
			string text = reader.ReadToEnd();
			Log.Info(text);
		}

		private static void TestTargetFrameworks()
		{
#if NET5_0_OR_GREATER
			Log.Info("NET 5.0 or greater Supported framworks");
#endif
#if NETCOREAPP3_0_OR_GREATER
			Log.Info("NETCOREAPP 3.0 or greater Supported framworks");
#endif
#if NETSTANDARD2_0_OR_GREATER
			Log.Info("NET Standard 1.1 or greater Supported framworks");
#endif
#if NETSTANDARD1_1_OR_GREATER
			Log.Info("NET Standard 1.1 or greater Supported framworks");
#endif
		}

		private static void TestTree()
		{
			DbxFolder folder1 = new (1, 0, "A", null);
			DbxFolder folder2 = new (2, 4, "B", null);
			DbxFolder folder3 = new (3, 0, "C", null);
			DbxFolder folder4 = new (4, 5, "D", null);
			DbxFolder folder5 = new (5, 0, "E", null);

			IList<DbxFolder> folders =
			[
				folder1,
				folder2,
				folder3,
				folder4,
				folder5
			];

			DbxFolder folder = new (0, 0, "root", null);

			folder.GetChildren(folders);

			IList<uint> orderedIndexes = [];
			folder.SetOrderedIndexes(orderedIndexes);
		}
	}
}
