/////////////////////////////////////////////////////////////////////////////
// <copyright file="Program.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Email.ToolKit;
using Microsoft.Office.Interop.Outlook;
using Serilog;
using Serilog.Configuration;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using CommonLogging = Common.Logging;

[assembly: CLSCompliant(true)]

namespace DigitalZenWorks.Email.ToolKit.Application
{
	/// <summary>
	/// Dbx to pst program class.
	/// </summary>
	public static class Program
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private static readonly string[] Commands =
		{
			"dbx-to-pst", "eml-to-pst", "help", "list-folders",
			"list-top-senders", "list-total-duplicates", "merge-folders",
			"merge-stores", "move-folder", "remove-duplicates",
			"remove-empty-folders"
		};

		/// <summary>
		/// The program's main entry point.
		/// </summary>
		/// <param name="arguments">The arguments given to the program.</param>
		/// <returns>A value indicating success or not.</returns>
		public static int Main(string[] arguments)
		{
			int result = -1;

			try
			{
				LogInitialization();
				string version = GetVersion();

				Log.Info("Starting DigitalZenWorks.Email.ToolKit Version: " +
					version);

				IList<Command> commands = new List<Command>();

				Command help = new ("help");
				help.Description = "Show this information";
				commands.Add(help);

				CommandOption encoding = new ("e", "encoding", true);
				IList<CommandOption> options = new List<CommandOption>();
				options.Add(encoding);

				Command dbxToPst = new (
					"dbx-to-pst", options, 1, "Migrate dbx files to pst file");
				commands.Add(dbxToPst);

				CommandOption adjust = new ("a", "adjust");
				options = new List<CommandOption>();
				options.Add(adjust);

				Command emlToPst = new (
					"eml-to-pst", options, 1, "Migrate eml files to pst file");
				commands.Add(dbxToPst);

				CommandOption recurse = new ("r", "recurse");
				options = new List<CommandOption>();
				options.Add(recurse);

				Command listFolders = new (
					"list-folders",
					options,
					1,
					"List all sub folders of a given store or folder");
				commands.Add(listFolders);

				Command listTopSenders = new (
					"list-top-senders",
					null,
					1,
					"List the top senders of a given store");
				commands.Add(listTopSenders);

				Command listTotalDuplicates = new (
					"list-total-duplicates",
					null,
					1,
					"List all duplicates in a given store");
				commands.Add(listTotalDuplicates);

				CommandOption dryRun = new ("n", "dryrun");
				options = new List<CommandOption>();
				options.Add(dryRun);

				Command mergeFolders = new (
					"merge-folders",
					options,
					1,
					"Merge duplicate Outlook folders");
				commands.Add(mergeFolders);

				Command mergeStores = new (
					"merge-stores", null, 2, "Merge one store into another");
				commands.Add(mergeStores);

				Command moveFolders = new (
					"move-folders", null, 4, "Move one folder to another");
				commands.Add(moveFolders);

				CommandOption flush = new ("s", "flush");
				options = new List<CommandOption>();
				options.Add(dryRun);
				options.Add(flush);

				Command removeDuplicates = new (
					"merge-folders",
					options,
					1,
					"Merge duplicate Outlook folders");
				commands.Add(removeDuplicates);

				Command removeEmptyFolders = new (
					"remove-empty-folders", null, 1, "Prune empty folders");
				commands.Add(removeEmptyFolders);

				CommandLineArguments commandLine = new (commands, arguments);

				if (commandLine.ValidArguments == false)
				{
					Log.Error(commandLine.ErrorMessage);

					commandLine.UsageStatement =
						"Det command <options> <path.to.source> <path.to.pst>";
					commandLine.ShowHelp();
				}
				else
				{

				}

				bool valid = ValidateArguments(arguments);

				if (arguments != null && valid == true)
				{
					string command = arguments[0];
					string pstLocation;

					Log.Info("Command is: " + command);

					for (int index = 1; index < arguments.Length; index++)
					{
						string message = string.Format(
							CultureInfo.InvariantCulture,
							"Parameter {0}: {1}",
							index.ToString(CultureInfo.InvariantCulture),
							arguments[index]);

						Log.Info(message);
					}

					switch (command)
					{
						case "dbx-to-pst":
							string dbxLocation = GetDbxLocation(arguments);

							pstLocation =
								GetPstLocation(arguments, dbxLocation, 2);

							result = DbxToPst(arguments, dbxLocation, pstLocation);
							break;
						case "eml-to-pst":
							string emlLocation = GetEmlLocation(arguments);
							int index = arguments.Length - 1;
							pstLocation =
								GetPstLocation(arguments, emlLocation, index);

							result =
								EmlToPst(arguments, emlLocation, pstLocation);
							break;
						case "help":
							ShowHelp();
							result = 0;
							break;
						case "list-folders":
							result = ListFolders(arguments);
							break;
						case "list-top-senders":
							result = ListTopSenders(arguments);
							break;
						case "list-total-duplicates":
							result = ListTotalDuplicates(arguments);
							break;
						case "merge-folders":
							result = MergeFolders(arguments);
							break;
						case "merge-stores":
							result = MergeStores(arguments);
							break;
						case "move-folder":
							MoveFolder(arguments);
							break;
						case "remove-duplicates":
							result = RemoveDuplicates(arguments);
							break;
						case "remove-empty-folders":
							result = RemoveEmptyFolders(arguments);
							break;
						default:
							result = ProcessDirect(arguments);
							break;
					}
				}
				else
				{
					Log.Error("Invalid arguments");

					ShowHelp();
				}
			}
			catch (System.Exception exception)
			{
				Log.Error(exception.ToString());

				throw;
			}

			return result;
		}

		private static int ArgumentsContainPstFile(string[] arguments)
		{
			int pstFileIndex = 0;

			if (arguments.Length > 1)
			{
				for (int index = 1; index < arguments.Length; index++)
				{
					string extension = Path.GetExtension(arguments[index]);

					if (extension.Equals(".pst", StringComparison.Ordinal))
					{
						pstFileIndex = index;
						break;
					}
				}
			}

			return pstFileIndex;
		}

		private static int DbxToPst(
			string[] arguments, string dbxLocation, string pstLocation)
		{
			int result = -1;
			Encoding encoding = GetEncoding(arguments);

			bool success = Migrate.DbxToPst(
				dbxLocation, pstLocation, encoding);

			if (success == true)
			{
				result = 0;
			}

			return result;
		}

		private static int EmlToPst(
			string[] arguments, string emlLocation, string pstLocation)
		{
			int result = -1;

			bool adjust = false;

			if (arguments.Contains("-a") ||
				arguments.Contains("--adjust"))
			{
				adjust = true;
			}

			bool success = Migrate.EmlToPst(emlLocation, pstLocation, adjust);

			if (success == true)
			{
				result = 0;
			}

			return result;
		}

		private static FileVersionInfo GetAssemblyInformation()
		{
			FileVersionInfo fileVersionInfo = null;

			Assembly assembly = Assembly.GetExecutingAssembly();

			string location = assembly.Location;

			if (string.IsNullOrWhiteSpace(location))
			{
				// Single file apps have no assemblies.
				Process process = Process.GetCurrentProcess();
				location = process.MainModule.FileName;
			}

			if (!string.IsNullOrWhiteSpace(location))
			{
				fileVersionInfo = FileVersionInfo.GetVersionInfo(location);
			}

			return fileVersionInfo;
		}

		private static string GetDbxLocation(string[] arguments)
		{
			string dbxLocation = null;

			for (int index = 1; index < arguments.Length; index++)
			{
				string argument = arguments[index];

				if (argument.Equals(
					"--encoding", StringComparison.OrdinalIgnoreCase) ||
					argument.Equals(
						"-e", StringComparison.OrdinalIgnoreCase))
				{
					index += 2;

					dbxLocation = arguments[index];
					break;
				}
			}

			return dbxLocation;
		}

		private static string GetEmlLocation(string[] arguments)
		{
			// skip pst path
			int index = arguments.Length - 2;

			string emlLocation = arguments[index];

			return emlLocation;
		}

		private static int GetCount(string[] arguments)
		{
			int count = 25;

			if (arguments.Contains("-c") ||
				arguments.Contains("--count"))
			{
				for (int index = 1; index < arguments.Length; index++)
				{
					string argument = arguments[index];

					if (argument.Equals(
						"--count", StringComparison.OrdinalIgnoreCase) ||
						argument.Equals(
							"-c", StringComparison.OrdinalIgnoreCase))
					{
						string rawCount = arguments[index + 1];
						count = Convert.ToInt32(
							rawCount, CultureInfo.InvariantCulture);

						break;
					}
				}
			}

			return count;
		}

		private static Encoding GetEncoding(string[] arguments)
		{
			Encoding encoding = null;

			if (arguments.Contains("-e") ||
				arguments.Contains("--encoding"))
			{
				for (int index = 1; index < arguments.Length; index++)
				{
					string argument = arguments[index];

					if (argument.Equals(
						"--encoding", StringComparison.OrdinalIgnoreCase) ||
						argument.Equals(
							"-e", StringComparison.OrdinalIgnoreCase))
					{
						string encodingName = arguments[index + 1];

						Encoding.RegisterProvider(
							CodePagesEncodingProvider.Instance);
						encoding = Encoding.GetEncoding(encodingName);

						break;
					}
				}
			}

			return encoding;
		}

		private static string GetFolderPath(string[] arguments)
		{
			string folderPath = null;

			for (int index = 1; index < arguments.Length; index++)
			{
				string argument = arguments[index];

				// if the file exists, it is the PST file argument
				bool fileExists = File.Exists(argument);

				if (fileExists == false && !argument.Equals(
					"--recurse", StringComparison.OrdinalIgnoreCase) &&
					!argument.Equals(
						"-r", StringComparison.OrdinalIgnoreCase))
				{
					folderPath = arguments[index];
					break;
				}
			}

			return folderPath;
		}

		private static string GetPstLocation(
			string[] arguments, string source, int index)
		{
			string pstLocation = null;

			if (arguments.Length > index)
			{
				pstLocation = arguments[index];
			}

			if (string.IsNullOrWhiteSpace(pstLocation))
			{
				if (Directory.Exists(source))
				{
					pstLocation = source + ".pst";
				}
				else if (File.Exists(source))
				{
					pstLocation =
						Path.ChangeExtension(source, ".pst");
				}
			}

			return pstLocation;
		}

		private static string GetVersion()
		{
			FileVersionInfo fileVersionInfo = GetAssemblyInformation();

			string assemblyVersion = fileVersionInfo.FileVersion;

			return assemblyVersion;
		}

		private static int ListFolders(string[] arguments)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			bool recurse = false;

			if (arguments.Contains("-r") || arguments.Contains("--recurse"))
			{
				recurse = true;
			}

			int pstFileIndex = ArgumentsContainPstFile(arguments);

			if (pstFileIndex > 0)
			{
				string pstFile = arguments[pstFileIndex];
				string folderPath = GetFolderPath(arguments);

				IList<string> folderNames =
					outlookStore.ListFolders(pstFile, folderPath, recurse);

				IOrderedEnumerable<string> sortedFolderName =
					folderNames.OrderBy(x => x);

				foreach (string folderName in sortedFolderName)
				{
					Console.WriteLine(folderName);
				}
			}

			return 0;
		}

		private static int ListTopSenders(string[] arguments)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			int pstFileIndex = ArgumentsContainPstFile(arguments);

			if (pstFileIndex > 0)
			{
				string pstFile = arguments[pstFileIndex];

				int count = GetCount(arguments);

				IList<KeyValuePair<string, int>> topSenders =
					outlookStore.ListTopSenders(pstFile, count);

				foreach (KeyValuePair<string, int> sender in topSenders)
				{
					string message = string.Format(
						CultureInfo.InvariantCulture,
						"{0}: {1}",
						sender.Key,
						sender.Value.ToString(CultureInfo.InvariantCulture));
					Console.WriteLine(message);
				}
			}

			return 0;
		}

		private static int ListTotalDuplicates(string[] arguments)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			int pstFileIndex = ArgumentsContainPstFile(arguments);

			if (pstFileIndex > 0)
			{
				string pstFilePath = arguments[pstFileIndex];

				IDictionary<string, IList<string>> duplicates =
					outlookStore.GetTotalDuplicates(pstFilePath);

				ListTotalDuplicatesOutput(duplicates, true);
				ListTotalDuplicatesOutput(duplicates, false);
			}

			return 0;
		}

		private static void ListTotalDuplicatesOutput(
			IDictionary<string, IList<string>> duplicates, bool useLog)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			bool duplicatesFound = false;

			foreach (KeyValuePair<string, IList<string>> item in
				duplicates)
			{
				IList<string> duplicateSet = item.Value;

				if (duplicateSet.Count > 1)
				{
					duplicatesFound = true;
					string entryId1 = duplicateSet[0];

					MailItem mailItem =
						outlookStore.GetMailItemFromEntryId(entryId1);

					string synopses =
						OutlookFolder.GetMailItemSynopses(mailItem);

					string message = string.Format(
						CultureInfo.InvariantCulture,
						"Duplicates Found for: {0}",
						synopses);

					if (useLog == true)
					{
						Log.Info(message);
					}
					else
					{
						Console.WriteLine(message);
					}

					foreach (string entryId in duplicateSet)
					{
						mailItem =
							outlookStore.GetMailItemFromEntryId(entryId);

						MAPIFolder parent = mailItem.Parent;
						string path = OutlookFolder.GetFolderPath(parent);

						message = "At: " + path;

						if (useLog == true)
						{
							Log.Info(message);
						}
						else
						{
							Console.WriteLine(message);
						}
					}
				}
			}

			if (duplicatesFound == false)
			{
				string message = "No duplicates found";

				if (useLog == true)
				{
					Log.Info(message);
				}
				else
				{
					Console.WriteLine(message);
				}
			}
		}

		private static void LogInitialization()
		{
			string baseDataDirectory = Environment.GetFolderPath(
				Environment.SpecialFolder.ApplicationData,
				Environment.SpecialFolderOption.Create);

			baseDataDirectory += @"\DigitalZenWorks\Email.Toolkit";
			string logFilePath = baseDataDirectory + @"\Email.Toolkit.log";

			string outputTemplate = "[{Timestamp:yyyy-MM-dd HH:mm:ss} " +
				"{Level:u3}] {Message:lj}{NewLine}{Exception}";

			LoggerConfiguration configuration = new ();
			LoggerSinkConfiguration sinkConfiguration = configuration.WriteTo;
			sinkConfiguration.Console(LogEventLevel.Verbose, outputTemplate);
			sinkConfiguration.File(
				logFilePath, LogEventLevel.Verbose, outputTemplate);
			Serilog.Log.Logger = configuration.CreateLogger();

			LogManager.Adapter =
				new CommonLogging.Serilog.SerilogFactoryAdapter();
		}

		private static int MergeFolders(string[] arguments)
		{
			bool dryRun = false;

			if (arguments.Contains("-n") ||
				arguments.Contains("--dryrun"))
			{
				dryRun = true;
			}

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			int pstFileIndex = ArgumentsContainPstFile(arguments);

			if (pstFileIndex > 0)
			{
				string pstFile = arguments[pstFileIndex];

				outlookStore.MergeFolders(pstFile, dryRun);
			}
			else
			{
				outlookAccount.MergeFolders(dryRun);
			}

			return 0;
		}

		private static int MergeStores(string[] arguments)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			if (arguments.Length > 2)
			{
				string sourcePst = arguments[1];
				string destinationPst = arguments[2];

				outlookStore.MergeStores(
					sourcePst, destinationPst);
			}

			return 0;
		}

		private static int MoveFolder(string[] arguments)
		{
			string sourcePst = arguments[1];
			string sourcePath = arguments[2];
			string destinationPst = arguments[3];
			string destinationPath = arguments[4];

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			outlookStore.MoveFolder(
				sourcePst,
				sourcePath,
				destinationPst,
				destinationPath);

			return 0;
		}

		private static int ProcessDirect(string[] arguments)
		{
			int result = -1;

			if (arguments.Length > 0)
			{
				string location = arguments[0];

				string pstLocation = GetPstLocation(arguments, location, 1);
				if (Directory.Exists(location))
				{
					result = ProcessDirectDirectory(
						arguments, location, pstLocation);
				}
				else if (File.Exists(location))
				{
					result =
						ProcessDirectFile(arguments, location, pstLocation);
				}
				else
				{
					string message =
						"Argument supplied is neither a directory nor a file.";
					Log.Error(message);
				}
			}

			return result;
		}

		private static int ProcessDirectDirectory(
			string[] arguments, string location, string pstLocation)
		{
			int result = -1;

			string[] files = Directory.GetFiles(location, "*.dbx");

			if (files.Length > 0)
			{
				result = DbxToPst(arguments, location, pstLocation);
			}
			else
			{
				IEnumerable<string> emlFiles = EmlMessages.GetFiles(location);

				if (emlFiles.Any())
				{
					result = EmlToPst(arguments, location, pstLocation);
				}
			}

			return result;
		}

		private static int ProcessDirectFile(
			string[] arguments, string location, string pstLocation)
		{
			int result = -1;

			string extension = Path.GetExtension(location);

			if (extension.Equals(".dbx", StringComparison.Ordinal))
			{
				result = DbxToPst(arguments, location, pstLocation);
			}
			else if (extension.Equals(".eml", StringComparison.Ordinal) ||
				extension.Equals(".txt", StringComparison.Ordinal))
			{
				result = EmlToPst(arguments, location, pstLocation);
			}

			return result;
		}

		private static int RemoveDuplicates(string[] arguments)
		{
			bool dryRun = false;
			bool flush = false;

			if (arguments.Contains("-n") ||
				arguments.Contains("--dryrun"))
			{
				dryRun = true;
			}

			if (arguments.Contains("-s") ||
				arguments.Contains("--flush"))
			{
				if (dryRun == true)
				{
					// Obviously, ignore flush if dryRun is set.
					Log.Warn("Ignoring flush option as dryRun is set");
				}
				else
				{
					flush = true;
				}
			}

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			int pstFileIndex = ArgumentsContainPstFile(arguments);

			if (pstFileIndex > 0)
			{
				string pstFile = arguments[pstFileIndex];

				outlookStore.RemoveDuplicates(
					pstFile, dryRun, flush);
			}
			else
			{
				outlookAccount.RemoveDuplicates(dryRun, flush);
			}

			return 0;
		}

		private static int RemoveEmptyFolders(string[] arguments)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;

			int pstFileIndex = ArgumentsContainPstFile(arguments);
			int removedFolders;

			if (pstFileIndex > 0)
			{
				OutlookStore outlookStore = new (outlookAccount);
				string pstFile = arguments[pstFileIndex];

				removedFolders = outlookStore.RemoveEmptyFolders(pstFile);
			}
			else
			{
				removedFolders = outlookAccount.RemoveEmptyFolders();
			}

			Console.WriteLine("Folder removed: " +
				removedFolders.ToString(CultureInfo.InvariantCulture));

			return 0;
		}

		private static void ShowHelp(string additionalMessage = null)
		{
			Assembly assembly = Assembly.GetExecutingAssembly();
			AssemblyName assemblyName = assembly.GetName();
			string name = assemblyName.Name;

			FileVersionInfo versionInfo = GetAssemblyInformation();
			string companyName = versionInfo.CompanyName;
			string copyright = versionInfo.LegalCopyright;
			string assemblyVersion = versionInfo.FileVersion;

			string header = string.Format(
				CultureInfo.CurrentCulture,
				"{0} {1} {2} {3}",
				name,
				assemblyVersion,
				copyright,
				companyName);
			Log.Info(header);

			if (!string.IsNullOrWhiteSpace(additionalMessage))
			{
				Log.Info(additionalMessage);
			}

			Log.Info("Usage:");
			Log.Info("DigitalZenWorks.Email.ToolKit & lt; " +
				"command <options> <path.to.source> <path.to.pst>");

			Log.Info("Commands:");
			Log.Info("dbx-to-pst             Migrate dbx files to pst file");
			Log.Info("eml-to-pst             Migrate eml files to pst file");
			Log.Info("list-folders           " +
				"List all sub folders of a given folder");
			Log.Info("list-top-senders       " +
				"List the top senders of a given store");
			Log.Info("list-total-duplicates  " +
				"List all duplicates in all folders in a given store");
			Log.Info("move-folder            " +
				"Move folder to a different location");
			Log.Info("merge-folders          Merge duplicate folders");
			Log.Info("merge-stores           Merge one store into another");
			Log.Info("remove-duplicates      Prune empty folders");
			Log.Info("remove-empty-folders   Prune empty folders");
			Log.Info("help                   Show this information");
		}

		private static bool ValidateArguments(string[] arguments)
		{
			bool valid = false;

			if (arguments != null && arguments.Length > 0)
			{
				string command = arguments[0];

				if (Commands.Contains(command))
				{
					if (command.Equals(
						"help", StringComparison.OrdinalIgnoreCase) ||
						command.Equals(
						"list-folders", StringComparison.OrdinalIgnoreCase) ||
						command.Equals(
						"list-top-senders",
						StringComparison.OrdinalIgnoreCase) ||
						command.Equals(
						"list-total-duplicates",
						StringComparison.OrdinalIgnoreCase) ||
						command.Equals(
						"merge-folders", StringComparison.OrdinalIgnoreCase) ||
						command.Equals(
						"remove-duplicates",
						StringComparison.OrdinalIgnoreCase) ||
						command.Equals(
						"remove-empty-folders",
						StringComparison.OrdinalIgnoreCase))
					{
						valid = true;
					}
					else if (command.Equals(
						"dbx-to-pst", StringComparison.OrdinalIgnoreCase))
					{
						string dbxLocation = GetDbxLocation(arguments);

						if (Directory.Exists(dbxLocation) ||
							File.Exists(dbxLocation))
						{
							valid = true;
						}
					}
					else if (command.Equals(
						"eml-to-pst", StringComparison.OrdinalIgnoreCase))
					{
						string emlLocation = GetEmlLocation(arguments);

						if (Directory.Exists(emlLocation) ||
							File.Exists(emlLocation))
						{
							valid = true;
						}
					}
					else if (command.Equals(
						"move-folder", StringComparison.OrdinalIgnoreCase))
					{
						if (arguments.Length > 4)
						{
							valid = true;
						}
					}
					else if (arguments.Length > 1)
					{
						if (arguments.Length > 2 || !command.Equals(
							"merge-stores", StringComparison.OrdinalIgnoreCase))
						{
							string location = arguments[1];

							if (Directory.Exists(location) ||
								File.Exists(location))
							{
								valid = true;
							}
						}
					}
				}
				else
				{
					if (File.Exists(command))
					{
						string extension = Path.GetExtension(command);

						// Command inferred from file type.
						if (extension.Equals(
							".dbx", StringComparison.OrdinalIgnoreCase) ||
							extension.Equals(
								".eml", StringComparison.OrdinalIgnoreCase) ||
							extension.Equals(
								".txt", StringComparison.OrdinalIgnoreCase))
						{
							valid = true;
						}
					}
					else if (Directory.Exists(command))
					{
						string[] dbxFiles =
							Directory.GetFiles(command, "*.dbx");
						string[] emlFiles =
							Directory.GetFiles(command, "*.eml");
						string[] txtFiles =
							Directory.GetFiles(command, "*.txt");

						if (dbxFiles.Length > 0 || emlFiles.Length > 0 ||
							txtFiles.Length > 0)
						{
							valid = true;
						}
					}
				}
			}

			return valid;
		}
	}
}
