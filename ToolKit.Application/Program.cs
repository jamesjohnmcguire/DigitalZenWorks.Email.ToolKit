﻿/////////////////////////////////////////////////////////////////////////////
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

				Log.Info("Starting DigitalZenWorks.Email.ToolKit Version: " +
					version);

				bool valid = ValidateLocationArguments(arguments);

				if (arguments != null && valid == true)
				{
					string pstLocation;

					switch (arguments[0])
					{
						case "dbx-to-pst":
							string dbxLocation = arguments[1];
							pstLocation =
								GetPstLocation(arguments, dbxLocation, 2);

							result = DbxToPst(dbxLocation, pstLocation);
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
						case "merge-folders":
							result = MergeFolders(arguments);
							break;
						case "merge-stores":
							result = MergeStores(arguments);
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

		private static int DbxToPst(string dbxLocation, string pstLocation)
		{
			int result = -1;

			bool success = Migrate.DbxToPst(dbxLocation, pstLocation);

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
			Assembly assembly = Assembly.GetExecutingAssembly();

			AssemblyName assemblyName = assembly.GetName();
			Version version = assemblyName.Version;
			string assemblyVersion = version.ToString();

			string location = assembly.Location;

			if (string.IsNullOrWhiteSpace(location))
			{
				// Single file apps have no assemblies.
				Process process = Process.GetCurrentProcess();
				location = process.MainModule.FileName;
			}

			if (!string.IsNullOrWhiteSpace(location))
			{
				FileVersionInfo fileVersionInfo =
					FileVersionInfo.GetVersionInfo(location);
				assemblyVersion = fileVersionInfo.FileVersion;
			}

			return assemblyVersion;
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
				new Common.Logging.Serilog.SerilogFactoryAdapter();
		}

		private static int MergeFolders(string[] arguments)
		{
			OutlookAccount outlookAccount = OutlookAccount.Instance;
			OutlookStore outlookStore = new (outlookAccount);

			int pstFileIndex = ArgumentsContainPstFile(arguments);

			if (pstFileIndex > 0)
			{
				string pstFile = arguments[pstFileIndex];

				outlookStore.MergeFolders(pstFile);
			}
			else
			{
				outlookAccount.MergeFolders();
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

		private static int ProcessDirect(string[] arguments)
		{
			int result = -1;

			if (arguments.Length > 0)
			{
				string location = arguments[0];

				string pstLocation = GetPstLocation(arguments, location, 1);
				if (Directory.Exists(location))
				{
					result = ProcessDirectDirectory(location, pstLocation);
				}
				else if (File.Exists(location))
				{
					result = ProcessDirectFile(location, pstLocation);
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
			string location, string pstLocation)
		{
			int result = -1;

			string[] files = Directory.GetFiles(location, "*.dbx");

			if (files.Length > 0)
			{
				result = DbxToPst(location, pstLocation);
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
			string location, string pstLocation)
		{
			int result = -1;

			string extension = Path.GetExtension(location);

			if (extension.Equals(".dbx", StringComparison.Ordinal))
			{
				result = DbxToPst(location, pstLocation);
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

			if (arguments.Contains("-s") ||
				arguments.Contains("--flush"))
			{
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

			if (pstFileIndex > 0)
			{
				OutlookStore outlookStore = new (outlookAccount);
				string pstFile = arguments[pstFileIndex];

				outlookStore.RemoveEmptyFolders(pstFile);
			}
			else
			{
				outlookAccount.RemoveEmptyFolders();
			}

			return 0;
		}

		private static void ShowHelp(string additionalMessage = null)
		{
			Assembly assembly = Assembly.GetExecutingAssembly();
			string location = assembly.Location;

			FileVersionInfo versionInfo =
				FileVersionInfo.GetVersionInfo(location);

			string companyName = versionInfo.CompanyName;
			string copyright = versionInfo.LegalCopyright;

			AssemblyName assemblyName = assembly.GetName();
			string name = assemblyName.Name;
			Version version = assemblyName.Version;
			string assemblyVersion = version.ToString();

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

		private static bool ValidateLocationArguments(string[] arguments)
		{
			bool valid = false;

			if (arguments != null && arguments.Length > 0)
			{
				string command = arguments[0];

				if (arguments.Length > 1)
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
				else
				{
					if (command.Equals(
						"help", StringComparison.OrdinalIgnoreCase) ||
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
					else
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
				}
			}

			if (valid == false)
			{
				Log.Error("Invalid arguments");

				ShowHelp();
			}

			return valid;
		}
	}
}
