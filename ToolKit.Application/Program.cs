/////////////////////////////////////////////////////////////////////////////
// <copyright file="Program.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Email.ToolKit;
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
			"dbx-to-pst", "eml-to-pst", "help", "list", "merge-folders",
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

				bool valid = ValidateArguments(arguments);

				if (arguments != null && valid == true)
				{
					string command = arguments[0];
					string pstLocation;

					Log.Info("Command is: " + command);

					switch (command)
					{
						case "dbx-to-pst":
							string dbxLocation = GetDbxLocation(arguments);

							pstLocation =
								GetPstLocation(arguments, dbxLocation, 2);

							result = DbxToPst(arguments, dbxLocation, pstLocation);
							break;
						case "eml-to-pst":
							string emlLocation = arguments[1];
							pstLocation =
								GetPstLocation(arguments, emlLocation, 2);

							result = EmlToPst(emlLocation, pstLocation);
							break;
						case "help":
							ShowHelp();
							result = 0;
							break;
						case "list":
							result = List(arguments);
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
			catch (Exception exception)
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

		private static int EmlToPst(string emlLocation, string pstLocation)
		{
			int result = -1;

			bool success = Migrate.EmlToPst(emlLocation, pstLocation);

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

		private static int List(string[] arguments)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			int pstFileIndex = ArgumentsContainPstFile(arguments);

			if (pstFileIndex > 0)
			{
				string pstFile = arguments[pstFileIndex];
				string folderPath = arguments[2];

				IList<string> folderNames =
					outlookStore.ListFolders(pstFile, folderPath);

				foreach (string folderName in folderNames)
				{
					Console.WriteLine(folderName);
				}
			}

			return 0;
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
					result = EmlToPst(location, pstLocation);
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
				result = EmlToPst(location, pstLocation);
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
			else if (arguments.Contains("-s") ||
				arguments.Contains("--flush"))
			{
				// Obviously, ignore flush if dryRun is set.
				flush = true;
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

			Log.Info("Remove empty folder complete - total folders removed:" +
				removedFolders);

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
				"command & gt; &lt; path & gt;");

			Log.Info("Commands:");
			Log.Info("dbx-to-pst            Migrate dbx files to pst file");
			Log.Info("eml-to-pst            Migrate eml files to pst file");
			Log.Info("merge-folders         Merge duplicate folders");
			Log.Info("merge-stores          Merge one store into another");
			Log.Info("remove-duplicates     Prune empty folders");
			Log.Info("remove-empty-folders  Prune empty folders");
			Log.Info("help                  Show this information");
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
						"list", StringComparison.OrdinalIgnoreCase) ||
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
