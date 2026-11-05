/////////////////////////////////////////////////////////////////////////////
// <copyright file="Program.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

[assembly: global::System.CLSCompliant(true)]

namespace DigitalZenWorks.Email.ToolKit.Application;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DigitalZenWorks.CommandLine.Commands;
using DigitalZenWorks.Common.VersionUtilities;
using global::Common.Logging;
using Serilog;
using Serilog.Configuration;
using Serilog.Events;

/// <summary>
/// Dbx to pst program class.
/// </summary>
internal static class Program
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

	/// <summary>
	/// The program's main entry point.
	/// </summary>
	/// <param name="arguments">The arguments given to the program.</param>
	/// <returns>A value indicating success or not.</returns>
	public static async Task<int> Main(string[] arguments)
	{
		int result = -1;

		try
		{
			LogInitialization();

			OutlookService outlook = new();
			result = await RunAsync(arguments, outlook).ConfigureAwait(false);
		}
		catch (System.Exception exception)
		{
			Log.Error(exception.ToString());

			throw;
		}

		return result;
	}

	/// <summary>
	/// Validates arguments and connects once before executing an Outlook
	/// command.
	/// </summary>
	/// <param name="arguments">The arguments given to the program.</param>
	/// <param name="outlook">The service used by the selected command.</param>
	/// <returns>Zero for success; a nonzero value for failure.</returns>
	internal static async Task<int> RunAsync(
		string[] arguments, IOutlookService outlook)
	{
		int result = -1;

		List<Command> commands = GetCommands();

		CommandLineInstance commandLine =
			new(commands, arguments, InferCommand);

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

			DisplayParameters(command, arguments);

			Encoding.RegisterProvider(
				CodePagesEncodingProvider.Instance);

			switch (command.Name)
			{
				case "add-pst":
					result = Commands.AddPst(command, outlook);
					break;
				case "dbx-to-pst":
					result = Commands.DbxToPst(command);
					break;
				case "details":
					result = Commands.Details(command, outlook);
					break;
				case "eml-to-pst":
					result = Commands.EmlToPst(command);
					break;
				case "list-folders":
					result = Commands.ListFolders(command, outlook);
					break;
				case "list-ids":
					result = Commands.ListIds(command, outlook);
					break;
				case "list-top-senders":
					result = Commands.ListTopSenders(command, outlook);
					break;
				case "list-total-duplicates":
					result = Commands.ListTotalDuplicates(command, outlook);
					break;
				case "merge-folders":
					result = await Commands.MergeFolders(command, outlook).
						ConfigureAwait(false);
					break;
				case "merge-stores":
					result = await Commands.MergeStores(command, outlook).
						ConfigureAwait(false);
					break;
				case "move-folder":
					result = await Commands.MoveFolder(command, outlook).
						ConfigureAwait(false);
					break;
				case "remove-duplicates":
					result = await Commands.RemoveDuplicates(command, outlook).
						ConfigureAwait(false);
					break;
				case "remove-empty-folders":
					result = await Commands.RemoveEmptyFolders(command).
						ConfigureAwait(false);
					break;
				default:
				case "help":
					string title = GetTitle();
					commandLine.UseLog = false;
					commandLine.ShowHelp(title);

					result = 0;
					break;
			}
		}

		return result;
	}

	private static void DisplayParameters(
		Command command, string[] arguments)
	{
		if (!command.Name.Equals(
			"help", StringComparison.OrdinalIgnoreCase))
		{
			string version = VersionSupport.GetVersion();

			string message = string.Format(
				CultureInfo.InvariantCulture,
				"Starting Det [{0}] Version: {1}",
				"DigitalZenWorks.Email.ToolKit",
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

	private static List<Command> GetCommands()
	{
		List<Command> commands = [];

		Command help = new("help");
		help.Description = "Show this information";
		commands.Add(help);

		Command addPst = new("add-pst", null, 1, "Add a PST file");
		commands.Add(addPst);

		CommandOption encoding = new("e", "encoding", true);
		List<CommandOption> options = [encoding];

		Command dbxToPst = new(
			"dbx-to-pst", options, 1, "Migrate dbx files to pst file");
		commands.Add(dbxToPst);

		CommandOption adjust = new("a", "adjust");

		CommandOption closeStore = new("c", "close-store");

		options = [adjust, closeStore];

		Command emlToPst = new(
			"eml-to-pst", options, 2, "Migrate eml files to pst file");
		commands.Add(emlToPst);

		CommandOption recurse = new("r", "recurse");
		options = [recurse];

		Command listFolders = new(
			"list-folders",
			options,
			1,
			"List all sub folders of a given store or folder");
		commands.Add(listFolders);

		Command listIds = new(
			"list-ids",
			null,
			1,
			"List all entry IDs of items in a given folder");
		commands.Add(listIds);

		CommandOption count = new("c", "count");
		options = [count];
		Command listTopSenders = new(
			"list-top-senders",
			options,
			1,
			"List the top senders of a given store");
		commands.Add(listTopSenders);

		Command listTotalDuplicates = new(
			"list-total-duplicates",
			null,
			1,
			"List all duplicates in a given store");
		commands.Add(listTotalDuplicates);

		CommandOption dryRun = new("n", "dryrun");
		options = [dryRun];

		Command mergeFolders = new(
			"merge-folders",
			options,
			0,
			"Merge duplicate Outlook folders");
		commands.Add(mergeFolders);

		Command mergeStores = new(
			"merge-stores", null, 2, "Merge one store into another");
		commands.Add(mergeStores);

		Command moveFolders = new(
			"move-folder", null, 4, "Move one folder to another");
		commands.Add(moveFolders);

		CommandOption flush = new("s", "flush");
		options = [dryRun, flush];

		Command removeDuplicates = new(
			"remove-duplicates",
			options,
			0,
			"Merge duplicate Outlook folders");
		commands.Add(removeDuplicates);

		Command removeEmptyFolders = new(
			"remove-empty-folders", null, 1, "Prune empty folders");
		commands.Add(removeEmptyFolders);

		Command details = new(
			"details", null, 2, "Show details of given item");
		commands.Add(details);

		return commands;
	}

	private static IEnumerable<string> GetEmlFiles(string location)
	{
		List<string> extensions = [".eml", ".txt"];
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

	private static string GetTitle()
	{
		Assembly assembly = Assembly.GetExecutingAssembly();
		AssemblyName assemblyName = assembly.GetName();
		string name = assemblyName.Name;

		FileVersionInfo versionInfo =
			VersionSupport.GetAssemblyInformation();
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

	private static void LogInitialization()
	{
		string baseDataDirectory = Environment.GetFolderPath(
			Environment.SpecialFolder.ApplicationData,
			Environment.SpecialFolderOption.Create);

		baseDataDirectory += @"\DigitalZenWorks\Email.Toolkit";
		string logFilePath = baseDataDirectory + @"\Email.Toolkit.log";

		string outputTemplate = "[{Timestamp:yyyy-MM-dd HH:mm:ss} " +
			"{Level:u3}] {Message:lj}{NewLine}{Exception}";

		LoggerConfiguration configuration = new();
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
			new global::Common.Logging.Serilog.SerilogFactoryAdapter();
	}
}
