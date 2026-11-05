/////////////////////////////////////////////////////////////////////////////
// <copyright file="Commands.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Application;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DigitalZenWorks.CommandLine.Commands;
using global::Common.Logging;
using Microsoft.Office.Interop.Outlook;

/// <summary>
/// Implements the Outlook command-line operations.
/// </summary>
/// <remarks>
/// The dispatcher validates command syntax. Each handler validates its
/// timeout option before attempting an Outlook connection.
/// </remarks>
internal static class Commands
{
	private static readonly ILog Log = LogManager.GetLogger(
		System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

	/// <summary>
	/// Opens or creates the PST store specified by the command.
	/// </summary>
	/// <param name="command">The command containing the PST file path.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero if the store is available; otherwise, -1.</returns>
	internal static int AddPst(Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			string pstLocation = command.Parameters[0];

			OutlookSession session = (OutlookSession)outlook.Session;
			Store store = session.GetStore(pstLocation);

			if (store != null)
			{
				result = 0;
			}
		}

		return result;
	}

	/// <summary>
	/// Migrates DBX email data to a PST file.
	/// </summary>
	/// <param name="command">The command containing the DBX source, optional
	/// PST destination, and optional source encoding.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero if migration succeeds; otherwise, -1.</returns>
	internal static int DbxToPst(
		Command command, IOutlookService outlook)
	{
		bool connected = OutlookConnect(command, outlook);

		if (connected == false)
		{
			return -1;
		}

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

	/// <summary>
	/// Displays details of an Outlook item identified by its entry ID.
	/// </summary>
	/// <param name="command">The command containing the PST file path and
	/// item entry ID.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero when the command completes.</returns>
	internal static int Details(Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			string pstFilePath = command.Parameters[0];
			string entryId = command.Parameters[1];

			outlookStore.Details(pstFilePath, entryId);

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Migrates EML email data to a PST file.
	/// </summary>
	/// <param name="command">The command containing the EML source, PST
	/// destination, and migration options.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero if migration succeeds; otherwise, -1.</returns>
	internal static int EmlToPst(
		Command command, IOutlookService outlook)
	{
		bool connected = OutlookConnect(command, outlook);

		if (connected == false)
		{
			return -1;
		}

		int result = -1;

		string emlLocation = command.Parameters[0];
		string pstLocation = emlLocation;

		if (command.Parameters.Count > 1)
		{
			pstLocation = command.Parameters[1];
		}

		bool adjust = command.DoesOptionExist("a", "adjust");
		bool closeStore = command.DoesOptionExist("c", "close-store");

		bool success =
			Migrate.EmlToPst(emlLocation, pstLocation, adjust, closeStore);

		if (success == true)
		{
			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Writes folder names in alphabetical order to the console.
	/// </summary>
	/// <param name="command">The command containing the PST file path,
	/// optional folder path, and recursion option.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero when the command completes.</returns>
	internal static int ListFolders(Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			bool recurse = command.DoesOptionExist("r", "recurse");

			string pstFilePath = command.Parameters[0];
			string folderPath = null;

			if (command.Parameters.Count > 1)
			{
				folderPath = command.Parameters[1];
				folderPath = OutlookFolder.NormalizePath(folderPath);
			}

			IList<string> folderNames =
				outlookStore.ListFolders(pstFilePath, folderPath, recurse);

			IOrderedEnumerable<string> sortedFolderName =
				folderNames.OrderBy(x => x);

			foreach (string folderName in sortedFolderName)
			{
				Console.WriteLine(folderName);
			}

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Writes item entry IDs from a store or folder to the console.
	/// </summary>
	/// <param name="command">The command containing the PST file path and
	/// optional folder path.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero when the command completes.</returns>
	internal static int ListIds(Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			string pstFilePath = command.Parameters[0];
			string folderPath = null;

			if (command.Parameters.Count > 1)
			{
				folderPath = command.Parameters[1];
				folderPath = OutlookFolder.NormalizePath(folderPath);
			}

			IList<string> entryIds =
				outlookStore.GetIds(pstFilePath, folderPath);

			foreach (string entryId in entryIds)
			{
				Console.WriteLine(entryId);
			}

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Writes a store's top senders and their message counts to the console.
	/// </summary>
	/// <param name="command">The command containing the PST file path and
	/// optional sender count, which defaults to 25.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero when the command completes.</returns>
	internal static int ListTopSenders(Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

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

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Reports a store's duplicate items to the log and console.
	/// </summary>
	/// <param name="command">The command containing the PST file path.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>Zero when the command completes.</returns>
	internal static int ListTotalDuplicates(
		Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			string pstFilePath = command.Parameters[0];

			IDictionary<string, IList<string>> duplicates =
				outlookStore.GetTotalDuplicates(pstFilePath);

			ListTotalDuplicatesOutput(duplicates, true, outlook);
			ListTotalDuplicatesOutput(duplicates, false, outlook);

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Reports duplicate item summaries and the folder paths of each copy.
	/// Reports when no group contains duplicate items.
	/// </summary>
	/// <param name="duplicates">The groups of matching item entry IDs.</param>
	/// <param name="useLog">True to write to the log; false to use the console.
	/// </param>
	/// <param name="outlook">The connected Outlook service.</param>
	internal static void ListTotalDuplicatesOutput(
		IDictionary<string, IList<string>> duplicates,
		bool useLog,
		IOutlookService outlook)
	{
		OutlookSession session = (OutlookSession)outlook.Session;
		OutlookStore outlookStore = new(session);

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

				OutlookItem outlookItem = new(mailItem);
				string synopses = outlookItem.Synopses;

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
			string message =
				global::ToolKit.Application.Resources.NO_DUPLICATES_FOUND;

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

	/// <summary>
	/// Merges duplicate folders in a specified store or all connected stores.
	/// </summary>
	/// <param name="command">The command containing an optional PST file path
	/// and the dry-run option.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>A task returning zero when the command completes.</returns>
	internal static async Task<int> MergeFolders(
		Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			bool dryRun = command.DoesOptionExist("n", "dryrun");

			if (command.Parameters.Count > 0)
			{
				string pstFile = command.Parameters[0];

				await outlookStore.MergeFoldersAsync(pstFile, dryRun).
					ConfigureAwait(false);
			}
			else
			{
				OutlookAccount outlookAccount = OutlookAccount.Instance;

				await outlookAccount.MergeFoldersAsync(dryRun).
					ConfigureAwait(false);
			}

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Merges a source PST store into a destination PST store.
	/// </summary>
	/// <param name="command">The command containing the source and destination
	/// PST file paths, in that order.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>A task returning zero when the command completes.</returns>
	internal static async Task<int> MergeStores(
		Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			string sourcePst = command.Parameters[0];
			string destinationPst = command.Parameters[1];

			await outlookStore.MergeStoresAsync(sourcePst, destinationPst).
				ConfigureAwait(false);

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Moves a folder between the specified Outlook stores and folder paths.
	/// </summary>
	/// <param name="command">The command containing the source PST path,
	/// source folder path, destination PST path, and destination folder path.
	/// </param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>A task returning zero when the command completes.</returns>
	internal static async Task<int> MoveFolder(
		Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			string sourcePst = command.Parameters[0];
			string sourcePath = command.Parameters[1];
			sourcePath = OutlookFolder.NormalizePath(sourcePath);

			string destinationPst = command.Parameters[2];
			string destinationPath = command.Parameters[3];
			destinationPath = OutlookFolder.NormalizePath(destinationPath);

			await outlookStore.MoveFolderAsync(
				sourcePst,
				sourcePath,
				destinationPst,
				destinationPath).ConfigureAwait(false);

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Removes duplicate items from a specified store or all connected stores.
	/// </summary>
	/// <param name="command">The command containing an optional PST file path
	/// and the dry-run and flush options. Dry-run takes precedence over flush.
	/// </param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>A task returning zero when the command completes.</returns>
	internal static async Task<int> RemoveDuplicates(
		Command command, IOutlookService outlook)
	{
		int result = -1;

		bool connected = OutlookConnect(command, outlook);

		if (connected == true)
		{
			OutlookSession session = (OutlookSession)outlook.Session;
			OutlookStore outlookStore = new(session);

			bool dryRun = command.DoesOptionExist("n", "dryrun");
			bool flush = command.DoesOptionExist("s", "flush");

			if (dryRun == true && flush == true)
			{
				// Obviously, ignore flush if dryRun is set.
				Log.Warn("Ignoring flush option as dryRun is set");
				flush = false;
			}

			if (command.Parameters.Count > 0)
			{
				string pstFilePath = command.Parameters[0];

				await outlookStore.RemoveDuplicatesAsync(
					pstFilePath, dryRun, flush).ConfigureAwait(false);
			}
			else
			{
				OutlookAccount outlookAccount = OutlookAccount.Instance;

				await outlookAccount.RemoveDuplicatesAsync(dryRun, flush).
					ConfigureAwait(false);
			}

			result = 0;
		}

		return result;
	}

	/// <summary>
	/// Removes empty folders and writes the number removed to the console.
	/// </summary>
	/// <param name="command">The command containing an optional PST file path.
	/// When omitted, all connected stores are processed.</param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>A task returning zero when the command completes.</returns>
	internal static async Task<int> RemoveEmptyFolders(
		Command command, IOutlookService outlook)
	{
		bool connected = OutlookConnect(command, outlook);

		if (connected == false)
		{
			return -1;
		}

		OutlookAccount outlookAccount = OutlookAccount.Instance;

		int removedFolders;

		if (command.Parameters.Count > 0)
		{
			OutlookStore outlookStore = new(outlookAccount);
			string pstFilePath = command.Parameters[0];

			removedFolders =
				await outlookStore.RemoveEmptyFoldersAsync(
					pstFilePath).ConfigureAwait(false);
		}
		else
		{
			removedFolders =
				await outlookAccount.RemoveEmptyFoldersAsync().
					ConfigureAwait(false);
		}

		Console.WriteLine("Folder removed: " +
			removedFolders.ToString(CultureInfo.InvariantCulture));

		return 0;
	}

	/// <summary>
	/// Resolves the optional source encoding supplied by the command.
	/// </summary>
	/// <param name="command">The command with the encoding option.</param>
	/// <returns>The requested encoding, or null when no option is supplied.
	/// </returns>
	private static Encoding GetEncoding(Command command)
	{
		Encoding encoding = null;

		CommandOption optionFound = command.GetOption("e", "encoding");

		if (optionFound != null)
		{
			string encodingName = optionFound.Parameter;

			encoding = Encoding.GetEncoding(encodingName);
		}

		return encoding;
	}

	private static int GetTimeOutSeconds(Command command)
	{
		int timeoutSeconds = 10;
		CommandOption optionFound = command.GetOption("t", "timeout");

		if (optionFound != null)
		{
			int maxTimeoutSeconds = int.MaxValue / 1000;

			bool validTimeout = int.TryParse(
				optionFound.Parameter,
				NumberStyles.Integer,
				CultureInfo.InvariantCulture,
				out timeoutSeconds);

			if (validTimeout == false || timeoutSeconds < 0 ||
				timeoutSeconds > maxTimeoutSeconds)
			{
				Log.Error("The --timeout (-t) value must be an integer " +
					"between 0 and 2147483 seconds.");

				timeoutSeconds = -1;
			}
		}

		return timeoutSeconds;
	}

	/// <summary>
	/// Validates the optional probe timeout and connects to Outlook.
	/// </summary>
	/// <param name="command">The command with the optional timeout value.
	/// </param>
	/// <param name="outlook">The service used to connect to Outlook.</param>
	/// <returns>True when the timeout is valid and connection succeeds.
	/// </returns>
	private static bool OutlookConnect(
		Command command, IOutlookService outlook)
	{
		bool connected = false;

		int timeoutSeconds = GetTimeOutSeconds(command);

		if (timeoutSeconds >= 0)
		{
			connected = outlook.Connect(timeoutSeconds);

			if (connected == false)
			{
				Log.Error("Unable to Continue");
			}
		}

		return connected;
	}
}
