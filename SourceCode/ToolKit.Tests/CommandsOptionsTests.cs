/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandsOptionsTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.IO;
using DigitalZenWorks.CommandLine.Commands;
using DigitalZenWorks.Email.ToolKit.Application;
using NUnit.Framework;

/// <summary>
/// Tests reusable application argument and option processing.
/// </summary>
internal sealed class CommandsOptionsTests
{
	/// <summary>
	/// Verifies folder extraction preserves both root and child names.
	/// </summary>
	[Test]
	public void GetFolderPathPreservesRootAndChildNames()
	{
		string[] arguments =
			["Employment.pst", "Employment\\Employment", "exports"];

		string folderPath = CommandsOptions.GetFolderPath(arguments);

		Assert.That(folderPath, Is.EqualTo("Employment/Employment"));
		Assert.That(arguments[1], Is.EqualTo("Employment\\Employment"));
	}

	/// <summary>
	/// Verifies output extraction selects the third argument without creating
	/// a directory.
	/// </summary>
	[Test]
	public void GetOutputDirectorySelectsAndNormalizesThirdArgument()
	{
		Guid identifier = Guid.NewGuid();
		string name = identifier.ToString("N");
		string[] arguments = ["Employment.pst", "Employment", name];

		string outputDirectory = CommandsOptions.GetOutputDirectory(arguments);

		Assert.That(Path.IsPathFullyQualified(outputDirectory), Is.True);
		Assert.That(Path.GetFileName(outputDirectory), Is.EqualTo(name));
		Assert.That(Directory.Exists(outputDirectory), Is.False);
		Assert.That(arguments[2], Is.EqualTo(name));
	}

	/// <summary>
	/// Verifies store extraction selects and normalizes the first argument.
	/// </summary>
	[Test]
	public void GetStorePathSelectsAndNormalizesFirstArgument()
	{
		string[] arguments = ["Employment.pst", "Employment", "exports"];

		string storePath = CommandsOptions.GetStorePath(arguments);

		Assert.That(Path.IsPathFullyQualified(storePath), Is.True);
		Assert.That(Path.GetFileName(storePath), Is.EqualTo("Employment.pst"));
		Assert.That(arguments[0], Is.EqualTo("Employment.pst"));
	}

	/// <summary>
	/// Verifies optional positional arguments use inclusive count bounds.
	/// </summary>
	/// <param name="count">The supplied positional count.</param>
	/// <param name="expected">Whether the count is allowed.</param>
	[TestCase(0, false)]
	[TestCase(1, true)]
	[TestCase(2, true)]
	[TestCase(3, false)]
	public void HasArgumentCountSupportsOptionalArguments(
		int count, bool expected)
	{
		string[] arguments = new string[count];

		bool valid = CommandsOptions.HasArgumentCount(arguments, 1, 2);

		Assert.That(valid, Is.EqualTo(expected));
	}

	/// <summary>
	/// Verifies extensions are checked without changing filename casing.
	/// </summary>
	/// <param name="path">The filesystem path.</param>
	/// <param name="expected">Whether its extension is .pst.</param>
	[TestCase("Employment.PST", true)]
	[TestCase("Employment.txt", false)]
	[TestCase("Employment", false)]
	public void HasFileExtensionIgnoresCasing(string path, bool expected)
	{
		bool valid = CommandsOptions.HasFileExtension(path, ".pst");

		Assert.That(valid, Is.EqualTo(expected));
	}

	/// <summary>
	/// Verifies null and blank arguments fail the shared value check.
	/// </summary>
	/// <param name="value">The invalid positional value.</param>
	[TestCase((string?)null)]
	[TestCase("")]
	[TestCase(" \t")]
	public void HaveArgumentValuesRejectsBlankValues(string? value)
	{
		string[] arguments = ["present", value!];

		bool valid = CommandsOptions.HaveArgumentValues(arguments);

		Assert.That(valid, Is.False);
	}

	/// <summary>
	/// Verifies output path validation permits a directory not yet created.
	/// </summary>
	[Test]
	public void IsValidPathAcceptsMissingDirectory()
	{
		Guid identifier = Guid.NewGuid();
		string name = identifier.ToString("N");
		string path = Path.Combine(Path.GetTempPath(), name);

		bool valid = CommandsOptions.IsValidPath(path);

		Assert.That(valid, Is.True);
		Assert.That(Directory.Exists(path), Is.False);
	}

	/// <summary>
	/// Verifies an existing uppercase PST filename is accepted.
	/// </summary>
	[Test]
	public void IsValidPstFileAcceptsExistingPstPath()
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Options.");

		try
		{
			string path = Path.Combine(directory.FullName, "Employment.PST");
			File.WriteAllText(
				path, "Path validation does not inspect contents.");

			bool valid = CommandsOptions.IsValidPstFile(path);

			Assert.That(valid, Is.True);
		}
		finally
		{
			directory.Delete(true);
		}
	}

	/// <summary>
	/// Verifies an existing file must have a PST extension.
	/// </summary>
	[Test]
	public void IsValidPstFileRejectsDifferentExtension()
	{
		string path = Path.GetTempFileName();

		try
		{
			bool valid = CommandsOptions.IsValidPstFile(path);

			Assert.That(valid, Is.False);
			Assert.That(File.Exists(path), Is.True);
		}
		finally
		{
			File.Delete(path);
		}
	}

	/// <summary>
	/// Verifies a directory with a PST suffix is not treated as a file.
	/// </summary>
	[Test]
	public void IsValidPstFileRejectsDirectory()
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Options.");

		try
		{
			string path = Path.Combine(directory.FullName, "folder.pst");
			Directory.CreateDirectory(path);

			bool valid = CommandsOptions.IsValidPstFile(path);

			Assert.That(valid, Is.False);
		}
		finally
		{
			directory.Delete(true);
		}
	}

	/// <summary>
	/// Verifies malformed source paths fail without throwing.
	/// </summary>
	/// <param name="path">The invalid source path.</param>
	[TestCase((string?)null)]
	[TestCase("")]
	[TestCase(" ")]
	[TestCase("bad\0path")]
	public void IsValidPstFileRejectsInvalidPath(string? path)
	{
		bool valid = CommandsOptions.IsValidPstFile(path);

		Assert.That(valid, Is.False);
	}

	/// <summary>
	/// Verifies a missing PST is not created by source validation.
	/// </summary>
	[Test]
	public void IsValidPstFileRejectsMissingFile()
	{
		Guid identifier = Guid.NewGuid();
		string name = identifier.ToString("N") + ".pst";
		string path = Path.Combine(Path.GetTempPath(), name);

		bool valid = CommandsOptions.IsValidPstFile(path);

		Assert.That(valid, Is.False);
		Assert.That(File.Exists(path), Is.False);
	}

	/// <summary>
	/// Verifies path normalization does not create the requested output.
	/// </summary>
	[Test]
	public void TryGetFullPathNormalizesWithoutCreatingOutput()
	{
		Guid identifier = Guid.NewGuid();
		string relative = identifier.ToString("N") + ".pst";

		bool valid = CommandsOptions.TryGetFullPath(relative, out var path);

		Assert.That(valid, Is.True);
		Assert.That(path, Is.Not.Null);
		Assert.That(Path.IsPathFullyQualified(path!), Is.True);
		Assert.That(path, Does.EndWith(relative));
		Assert.That(File.Exists(path), Is.False);
	}

	/// <summary>
	/// Verifies invalid path values are rejected without throwing.
	/// </summary>
	/// <param name="value">The invalid filesystem path.</param>
	[TestCase((string?)null)]
	[TestCase("")]
	[TestCase(" ")]
	[TestCase("bad\0path")]
	public void TryGetFullPathRejectsInvalidValues(string? value)
	{
		bool valid = CommandsOptions.TryGetFullPath(value, out var path);

		Assert.That(valid, Is.False);
		Assert.That(path, Is.Null);
	}

	/// <summary>
	/// Verifies the callback-shaped validator uses supplied positional values
	/// without mutating them or requiring a second command-line parse.
	/// </summary>
	[Test]
	public void ValidateMailBodyExportAcceptsCommandAndPositionalValues()
	{
		string store = Path.GetTempFileName();
		string pst = Path.ChangeExtension(store, ".pst");
		File.Move(store, pst);
		try
		{
			Command command = new("get-mail-bodies", [], []);
			string[] arguments =
				[pst, "Employment\\Employment", Path.GetTempPath()];

			bool valid = CommandsOptions.ValidateMailBodyExport(
				command, arguments, out var error);

			Assert.That(valid, Is.True);
			Assert.That(error, Is.Null);
			Assert.That(arguments[1], Is.EqualTo("Employment\\Employment"));
			Assert.That(command.Parameters, Is.Empty);
		}
		finally
		{
			File.Delete(pst);
		}
	}

	/// <summary>
	/// Verifies combined source syntax does not replace separate arguments.
	/// </summary>
	[Test]
	public void ValidateMailBodyExportRejectsCombinedSource()
	{
		Command command = new("get-mail-bodies", [], []);
		string[] arguments = ["archive.pst::Employment", "output"];

		bool valid = CommandsOptions.ValidateMailBodyExport(
			command, arguments, out var error);

		Assert.That(valid, Is.False);
		Assert.That(error, Does.StartWith("Usage: "));
	}

	/// <summary>
	/// Verifies missing, blank, and excess arguments fail validation.
	/// </summary>
	[Test]
	public void ValidateMailBodyExportRejectsIncompleteParameters()
	{
		Command command = new("get-mail-bodies", [], []);
		string[][] invalid =
		[
			Array.Empty<string>(),
			["archive.pst"],
			["archive.pst", "Employment"],
			["archive.pst", " ", "output"],
			["archive.pst", "Employment", " "],
			["archive.pst", "Employment", "output", "extra"]
		];

		foreach (string[] arguments in invalid)
		{
			bool valid = CommandsOptions.ValidateMailBodyExport(
				command, arguments, out var error);
			Assert.That(valid, Is.False);
			Assert.That(error, Is.Not.Null);
		}
	}

	/// <summary>
	/// Verifies an invalid output path receives a specific validation error.
	/// </summary>
	[Test]
	public void ValidateMailBodyExportRejectsInvalidOutputPath()
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Options.");

		try
		{
			string path = Path.Combine(directory.FullName, "Employment.pst");
			File.WriteAllText(path, "Source path fixture.");
			Command command = new("get-mail-bodies", [], []);
			string[] arguments = [path, "Employment", "bad\0path"];

			bool valid = CommandsOptions.ValidateMailBodyExport(
				command, arguments, out var error);

			Assert.That(valid, Is.False);
			Assert.That(
				error, Does.StartWith("Invalid output directory path:"));
		}
		finally
		{
			directory.Delete(true);
		}
	}

	/// <summary>
	/// Verifies missing source files return a validation error.
	/// </summary>
	[Test]
	public void ValidateMailBodyExportRejectsMissingSource()
	{
		Command command = new("get-mail-bodies", [], []);
		Guid identifier = Guid.NewGuid();
		string name = identifier.ToString("N") + ".pst";
		string store = Path.Combine(Path.GetTempPath(), name);
		string[] arguments = [store, "Employment", Path.GetTempPath()];

		bool valid = CommandsOptions.ValidateMailBodyExport(
			command, arguments, out var error);

		Assert.That(valid, Is.False);
		Assert.That(error, Does.Contain(store));
	}

	/// <summary>
	/// Verifies usage errors use the supplied command's name.
	/// </summary>
	[Test]
	public void ValidateMailBodyExportReportsCommandUsage()
	{
		Command command = new("body-export-alias", [], []);

		bool valid = CommandsOptions.ValidateMailBodyExport(
			command, [], out var error);

		Assert.That(valid, Is.False);
		Assert.That(error, Does.StartWith("Usage: body-export-alias "));
	}
}
