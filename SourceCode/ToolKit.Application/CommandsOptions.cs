/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandsOptions.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Application;

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using DigitalZenWorks.CommandLine.Commands;

/// <summary>
/// Validates and reads arguments used by the EML export command.
/// </summary>
internal static class CommandsOptions
{
	/// <summary>
	/// Gets the Outlook folder path from validated positional arguments.
	/// </summary>
	/// <param name="arguments">The validated positional arguments.</param>
	/// <returns>The folder path with normalized separators.</returns>
	internal static string GetFolderPath(IList<string> arguments)
	{
		return OutlookFolder.NormalizePath(arguments[1]);
	}

	/// <summary>
	/// Gets the output directory from validated positional arguments.
	/// </summary>
	/// <param name="arguments">The validated positional arguments.</param>
	/// <returns>The absolute output directory path.</returns>
	internal static string GetOutputDirectory(IList<string> arguments)
	{
		return Path.GetFullPath(arguments[2]);
	}

	/// <summary>
	/// Gets the PST file path from validated positional arguments.
	/// </summary>
	/// <param name="arguments">The validated positional arguments.</param>
	/// <returns>The absolute PST file path.</returns>
	internal static string GetStorePath(IList<string> arguments)
	{
		return Path.GetFullPath(arguments[0]);
	}

	/// <summary>
	/// Checks an argument count against an inclusive range.
	/// </summary>
	/// <param name="arguments">The raw positional values.</param>
	/// <param name="minimum">The minimum allowed count.</param>
	/// <param name="maximum">The maximum allowed count.</param>
	/// <returns>True when the count is within the allowed range.</returns>
	internal static bool HasArgumentCount(
		IList<string> arguments, int minimum, int maximum)
	{
		bool isValid = false;

		if (arguments.Count >= minimum && arguments.Count <= maximum)
		{
			isValid = true;
		}

		return isValid;
	}

	/// <summary>
	/// Checks a filesystem extension without regard to casing.
	/// </summary>
	/// <param name="path">The filesystem path.</param>
	/// <param name="extension">The expected extension, including its
	/// dot.</param>
	/// <returns>True when the extension matches.</returns>
	internal static bool HasFileExtension(string path, string extension)
	{
		bool hasFileExtension = false;
		string actual = Path.GetExtension(path);

		if (actual.Equals(extension, StringComparison.OrdinalIgnoreCase))
		{
			hasFileExtension = true;
		}

		return hasFileExtension;
	}

	/// <summary>
	/// Checks that every supplied positional value contains nonblank text.
	/// </summary>
	/// <param name="arguments">The raw positional values.</param>
	/// <returns>True when no supplied value is null, empty, or whitespace.
	/// </returns>
	internal static bool HaveArgumentValues(IEnumerable<string> arguments)
	{
		bool result = true;

		foreach (string argument in arguments)
		{
			if (string.IsNullOrWhiteSpace(argument) == true)
			{
				result = false;
				break;
			}
		}

		return result;
	}

	/// <summary>
	/// Checks filesystem path syntax without creating an output directory.
	/// </summary>
	/// <param name="path">The path to check.</param>
	/// <returns>True when the path can be normalized.</returns>
	internal static bool IsValidPath(string? path)
	{
		bool result = TryGetFullPath(path, out _);

		return result;
	}

	/// <summary>
	/// Checks that a path identifies an existing file with a .pst extension.
	/// </summary>
	/// <param name="path">The source PST path.</param>
	/// <returns>True for an existing .pst file path.</returns>
	/// <remarks>This checks the path, not the file's internal PST format.
	/// </remarks>
	internal static bool IsValidPstFile(string? path)
	{
		bool result = false;

		if (TryGetFullPath(path, out string? fullPath) == true)
		{
			bool pstExtension = HasFileExtension(fullPath, ".pst");
			bool fileExists = File.Exists(fullPath);

			if (pstExtension == true && fileExists == true)
			{
				result = true;
			}
		}

		return result;
	}

	/// <summary>
	/// Normalizes a filesystem path without creating a file or directory.
	/// </summary>
	/// <param name="path">The raw filesystem path.</param>
	/// <param name="fullPath">The absolute path on success, or null.</param>
	/// <returns>True when normalization succeeds.</returns>
	internal static bool TryGetFullPath(
		string? path, [NotNullWhen(true)] out string? fullPath)
	{
		bool result = false;
		fullPath = null;

		if (string.IsNullOrWhiteSpace(path) == false)
		{
			try
			{
				fullPath = Path.GetFullPath(path);
				result = true;
			}
			catch (Exception exception) when
			(
				exception is ArgumentException ||
				exception is NotSupportedException ||
				exception is PathTooLongException)
			{
				result = false;
			}
		}

		return result;
	}

	/// <summary>
	/// Validates EML export arguments before any Outlook connection.
	/// </summary>
	/// <param name="command">The command metadata and parsed options.</param>
	/// <param name="arguments">The unmodified positional arguments supplied by
	/// CommandLine.Commands, without command names or option tokens.</param>
	/// <param name="errorMessage">The validation error, or null on success.
	/// </param>
	/// <returns>True when export arguments identify an existing PST.</returns>
	internal static bool ValidateEmlExport(
		Command command,
		IList<string> arguments,
		out string? errorMessage)
	{
		bool result = false;
		errorMessage = "Usage: " + command.Name + " <PST> <folder> <output>.";
		bool countValid = HasArgumentCount(arguments, 3, 3);
		bool valuesSupplied = HaveArgumentValues(arguments);

		if (countValid == true && valuesSupplied == true)
		{
			bool pstFileValid = IsValidPstFile(arguments[0]);
			bool outputPathValid = IsValidPath(arguments[2]);

			if (pstFileValid == false)
			{
				errorMessage = "Invalid or missing source PST: " +
					arguments[0];
			}
			else if (outputPathValid == false)
			{
				errorMessage = "Invalid output directory path: " +
					arguments[2];
			}
			else
			{
				errorMessage = null;
				result = true;
			}
		}

		return result;
	}
}
