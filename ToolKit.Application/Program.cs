/////////////////////////////////////////////////////////////////////////////
// <copyright file="Program.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.CommandLine.Commands;
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

				IList<Command> commands = GetCommands();

				CommandLineArguments commandLine =
					new (commands, arguments, InferCommand);

				commandLine.UseLog = true;
				commandLine.UsageStatement =
					"Det command <options> <path.to.source> <path.to.pst>";

				if (commandLine.ValidArguments == false)
				{
					Log.Error(commandLine.ErrorMessage);

					commandLine.ShowHelp();
				}
				else
				{
					Command command = commandLine.Command;

#pragma warning disable CA1062
					DisplayParameters(command, arguments);
#pragma warning restore CA1062

					switch (command.Name)
					{
						case "dbx-to-pst":
							result = DbxToPst(command);
							break;
						case "eml-to-pst":
							result = EmlToPst(command);
							break;
						case "list-folders":
							result = ListFolders(command);
							break;
						case "list-top-senders":
							result = ListTopSenders(command);
							break;
						case "list-total-duplicates":
							result = ListTotalDuplicates(command);
							break;
						case "merge-folders":
							result = MergeFolders(command);
							break;
						case "merge-stores":
							result = MergeStores(command);
							break;
						case "move-folder":
							MoveFolder(command);
							break;
						case "remove-duplicates":
							result = RemoveDuplicates(command);
							break;
						case "remove-empty-folders":
							result = RemoveEmptyFolders(command);
							break;
						default:
						case "help":
							string title = GetTitle();
							commandLine.ShowHelp(title);
							result = 0;
							break;
					}
				}
			}
			catch (System.Exception exception)
			{
				Log.Error(exception.ToString());

				throw;
			}

			return result;
		}

		private static void DisplayParameters(
			Command command, string[] arguments)
		{
			if (!command.Name.Equals(
				"help", StringComparison.OrdinalIgnoreCase))
			{
				string version = GetVersion();

				string message = string.Format(
					CultureInfo.InvariantCulture,
					"Starting Det ({0}) Version: {1}",
					"Starting DigitalZenWorks.Email.ToolKit",
					version);
				Log.Info(message);

				Log.Info("Command is: " + command.Name);

				for (int index = 1; index < arguments.Length; index++)
				{
					message = string.Format(
						CultureInfo.InvariantCulture,
						"Parameter {0}: {1}",
						index.ToString(CultureInfo.InvariantCulture),
						arguments[index]);

					Log.Info(message);
				}
			}
		}

		private static int DbxToPst(Command command)
		{
			int result = -1;
			Encoding encoding = GetEncoding(command);

			string dbxLocation = command.Parameters[0];
			string pstLocation = dbxLocation;

			if (command.Parameters.Count > 1)
			{
				pstLocation = command.Parameters[1];
			}

			bool success =
				Migrate.DbxToPst(dbxLocation, pstLocation, encoding);

			if (success == true)
			{
				result = 0;
			}

			return result;
		}

		private static int EmlToPst(Command command)
		{
			int result = -1;

			string emlLocation = command.Parameters[0];
			string pstLocation = emlLocation;

			if (command.Parameters.Count > 1)
			{
				pstLocation = command.Parameters[1];
			}

			bool adjust = command.DoesOptionExist("a", "adjust");

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

		private static IList<Command> GetCommands()
		{
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
			commands.Add(emlToPst);

			CommandOption recurse = new ("r", "recurse");
			options = new List<CommandOption>();
			options.Add(recurse);

			Command listFolders = new (
				"list-folders",
				options,
				1,
				"List all sub folders of a given store or folder");
			commands.Add(listFolders);

			CommandOption count = new ("c", "count");
			options = new List<CommandOption>();
			options.Add(count);
			Command listTopSenders = new (
				"list-top-senders",
				options,
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
				0,
				"Merge duplicate Outlook folders");
			commands.Add(mergeFolders);

			Command mergeStores = new (
				"merge-stores", null, 2, "Merge one store into another");
			commands.Add(mergeStores);

			Command moveFolders = new (
				"move-folder", null, 4, "Move one folder to another");
			commands.Add(moveFolders);

			CommandOption flush = new ("s", "flush");
			options = new List<CommandOption>();
			options.Add(dryRun);
			options.Add(flush);

			Command removeDuplicates = new (
				"remove-duplicates",
				options,
				0,
				"Merge duplicate Outlook folders");
			commands.Add(removeDuplicates);

			Command removeEmptyFolders = new (
				"remove-empty-folders", null, 1, "Prune empty folders");
			commands.Add(removeEmptyFolders);

			return commands;
		}

		private static IEnumerable<string> GetEmlFiles(string location)
		{
			List<string> extensions = new () { ".eml", ".txt" };
			IEnumerable<string> allFiles =
				Directory.EnumerateFiles(location, "*.*");

			IEnumerable<string> query =
				allFiles.Where(file =>
					file.EndsWith(
						extensions[0], StringComparison.OrdinalIgnoreCase) ||
					file.EndsWith(
						extensions[1], StringComparison.OrdinalIgnoreCase));

			return query;
		}

		private static Encoding GetEncoding(Command command)
		{
			Encoding encoding = null;

			CommandOption optionFound = command.GetOption("e", "encoding");

			if (optionFound != null)
			{
				string encodingName = optionFound.Parameter;

				Encoding.RegisterProvider(
					CodePagesEncodingProvider.Instance);
				encoding = Encoding.GetEncoding(encodingName);
			}

			return encoding;
		}

		private static string GetTitle()
		{
			Assembly assembly = Assembly.GetExecutingAssembly();
			AssemblyName assemblyName = assembly.GetName();
			string name = assemblyName.Name;

			FileVersionInfo versionInfo = GetAssemblyInformation();
			string companyName = versionInfo.CompanyName;
			string copyright = versionInfo.LegalCopyright;
			string assemblyVersion = versionInfo.FileVersion;

			string title = string.Format(
				CultureInfo.CurrentCulture,
				"{0} {1} {2} {3}",
				name,
				assemblyVersion,
				copyright,
				companyName);

			return title;
		}

		private static string GetVersion()
		{
			FileVersionInfo fileVersionInfo = GetAssemblyInformation();

			string assemblyVersion = fileVersionInfo.FileVersion;

			return assemblyVersion;
		}

		private static Command InferCommand(
			string argument, IList<Command> commands)
		{
			Command inferredCommand = null;

			if (Directory.Exists(argument))
			{
				string[] files = Directory.GetFiles(argument, "*.dbx");

				if (files.Length > 0)
				{
					inferredCommand = commands.SingleOrDefault(
						command => command.Name == "dbx-to-pst");
				}
				else
				{
					IEnumerable<string> emlFiles = GetEmlFiles(argument);

					if (emlFiles.Any())
					{
						inferredCommand = commands.SingleOrDefault(
							command => command.Name == "eml-to-pst");
					}
				}
			}
			else if (File.Exists(argument))
			{
				string extension = Path.GetExtension(argument);

				if (extension.Equals(".dbx", StringComparison.Ordinal))
				{
					inferredCommand = commands.SingleOrDefault(
						command => command.Name == "dbx-to-pst");
				}
				else if (extension.Equals(".eml", StringComparison.Ordinal) ||
					extension.Equals(".txt", StringComparison.Ordinal))
				{
					inferredCommand = commands.SingleOrDefault(
						command => command.Name == "eml-to-pst");
				}
			}

			return inferredCommand;
		}

		private static int ListFolders(Command command)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			bool recurse = command.DoesOptionExist("r", "recurse");

			string pstFilePath = command.Parameters[0];
			string folderPath = null;

			if (command.Parameters.Count > 1)
			{
				folderPath = command.Parameters[1];
			}

			IList<string> folderNames =
				outlookStore.ListFolders(pstFilePath, folderPath, recurse);

			IOrderedEnumerable<string> sortedFolderName =
				folderNames.OrderBy(x => x);

			foreach (string folderName in sortedFolderName)
			{
				Console.WriteLine(folderName);
			}

			return 0;
		}

		private static int ListTopSenders(Command command)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			string pstFilePath = command.Parameters[0];
			int count = 25;

			CommandOption optionFound = command.GetOption("c", "count");

			if (optionFound != null)
			{
				count = Convert.ToInt32(
					optionFound.Parameter, CultureInfo.InvariantCulture);
			}

			IList<KeyValuePair<string, int>> topSenders =
				outlookStore.ListTopSenders(pstFilePath, count);

			foreach (KeyValuePair<string, int> sender in topSenders)
			{
				string message = string.Format(
					CultureInfo.InvariantCulture,
					"{0}: {1}",
					sender.Key,
					sender.Value.ToString(CultureInfo.InvariantCulture));
				Console.WriteLine(message);
			}

			return 0;
		}

		private static int ListTotalDuplicates(Command command)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			string pstFilePath = command.Parameters[0];

			IDictionary<string, IList<string>> duplicates =
				outlookStore.GetTotalDuplicates(pstFilePath);

			ListTotalDuplicatesOutput(duplicates, true);
			ListTotalDuplicatesOutput(duplicates, false);

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

		private static int MergeFolders(Command command)
		{
			bool dryRun = command.DoesOptionExist("n", "dryrun");

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			if (command.Parameters.Count > 0)
			{
				string pstFile = command.Parameters[0];

				outlookStore.MergeFolders(pstFile, dryRun);
			}
			else
			{
				outlookAccount.MergeFolders(dryRun);
			}

			return 0;
		}

		private static int MergeStores(Command command)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			string sourcePst = command.Parameters[0];
			string destinationPst = command.Parameters[1];

			outlookStore.MergeStores(sourcePst, destinationPst);

			return 0;
		}

		private static int MoveFolder(Command command)
		{
			string sourcePst = command.Parameters[0];
			string sourcePath = command.Parameters[1];

			string destinationPst = command.Parameters[2];
			string destinationPath = command.Parameters[3];

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			outlookStore.MoveFolder(
				sourcePst,
				sourcePath,
				destinationPst,
				destinationPath);

			return 0;
		}

		private static int RemoveDuplicates(Command command)
		{
			bool dryRun = command.DoesOptionExist("n", "dryrun");
			bool flush = command.DoesOptionExist("s", "flush");

			if (dryRun == true && flush == true)
			{
				// Obviously, ignore flush if dryRun is set.
				Log.Warn("Ignoring flush option as dryRun is set");
				flush = false;
			}

			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			if (command.Parameters.Count > 0)
			{
				string pstFilePath = command.Parameters[0];

				outlookStore.RemoveDuplicates(pstFilePath, dryRun, flush);
			}
			else
			{
				outlookAccount.RemoveDuplicates(dryRun, flush);
			}

			return 0;
		}

		private static int RemoveEmptyFolders(Command command)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;

			int removedFolders;

			if (command.Parameters.Count > 0)
			{
				OutlookStore outlookStore = new (outlookAccount);
				string pstFilePath = command.Parameters[0];

				removedFolders = outlookStore.RemoveEmptyFolders(pstFilePath);
			}
			else
			{
				removedFolders = outlookAccount.RemoveEmptyFolders();
			}

			Console.WriteLine("Folder removed: " +
				removedFolders.ToString(CultureInfo.InvariantCulture));

			return 0;
		}
	}
}
