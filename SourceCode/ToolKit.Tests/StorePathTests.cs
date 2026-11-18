/////////////////////////////////////////////////////////////////////////////
// <copyright file="StorePathTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.IO;
using NUnit.Framework;

/// <summary>
/// Tests production path comparison using real files without Outlook.
/// </summary>
internal sealed class StorePathTests
{
	/// <summary>
	/// Verifies different files do not match even with identical content.
	/// </summary>
	[Test]
	public void AreEqualDistinguishesDifferentFiles()
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Path.");

		try
		{
			string first = Path.Combine(directory.FullName, "first.pst");
			string second = Path.Combine(directory.FullName, "second.pst");
			File.WriteAllText(first, "Same content.");
			File.WriteAllText(second, "Same content.");

			Assert.That(StorePath.AreEqual(first, second), Is.False);
		}
		finally
		{
			directory.Delete(true);
		}
	}

	/// <summary>
	/// Verifies aliases are resolved without requesting access to file data.
	/// </summary>
	[Test]
	public void AreEqualHandlesExclusivelyLockedFile()
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Path.");

		try
		{
			string path = Path.Combine(directory.FullName, "locked.pst");
			File.WriteAllText(path, "Existing file.");
			using DirectoryJunction junction = new(directory.FullName);
			string alias = Path.Combine(junction.Path, "locked.pst");
			using FileStream locked = new(
				path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

			Assert.That(StorePath.AreEqual(path, alias), Is.True);
		}
		finally
		{
			directory.Delete(true);
		}
	}

	/// <summary>
	/// Verifies case-only comparisons retain their existing behavior.
	/// </summary>
	[Test]
	public void AreEqualIgnoresPathCasing()
	{
		string path = Path.Combine(Path.GetTempPath(), "Archive.pst");

		Assert.That(StorePath.AreEqual(path, path.ToUpperInvariant()), Is.True);
	}

	/// <summary>
	/// Verifies a link in an ancestor directory resolves in either direction.
	/// </summary>
	/// <param name="reverse">Whether the linked path is first.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void AreEqualResolvesAncestorDirectoryLink(bool reverse)
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Path.");

		try
		{
			string nested = Path.Combine(directory.FullName, "nested");
			Directory.CreateDirectory(nested);
			string path = Path.Combine(nested, "archive.pst");
			File.WriteAllText(path, "Existing file.");
			using DirectoryJunction junction = new(directory.FullName);
			string alias = Path.Combine(junction.Path, "nested", "archive.pst");

			bool result;
			if (reverse)
			{
				result = StorePath.AreEqual(alias, path);
			}
			else
			{
				result = StorePath.AreEqual(path, alias);
			}

			Assert.That(result, Is.True);
			Assert.That(File.ReadAllText(path), Is.EqualTo("Existing file."));
		}
		finally
		{
			directory.Delete(true);
		}
	}

	/// <summary>
	/// Verifies resolved paths exceeding the initial native buffer are
	/// complete.
	/// </summary>
	[Test]
	public void AreEqualResolvesLongUnicodePath()
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Path.");

		try
		{
			string nested = directory.FullName;
			for (int index = 0; index < 3; index++)
			{
				nested = Path.Combine(nested, new string('a', 90));
			}

			Directory.CreateDirectory(nested);
			string path = Path.Combine(nested, "café.pst");
			File.WriteAllText(path, "Existing file.");
			using DirectoryJunction junction = new(nested);
			string alias = Path.Combine(junction.Path, "café.pst");
			string extendedPath = @"\\?\" + path;

			Assert.That(StorePath.AreEqual(extendedPath, alias), Is.True);
		}
		finally
		{
			directory.Delete(true);
		}
	}

	/// <summary>
	/// Verifies unavailable paths do not match or create missing files.
	/// </summary>
	/// <param name="existingFile">Whether one path names an existing
	/// file.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void AreEqualReturnsFalseForMissingFiles(bool existingFile)
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Path.");

		try
		{
			string path = Path.Combine(directory.FullName, "archive.pst");
			if (existingFile)
			{
				File.WriteAllText(path, "Existing file.");
			}

			using DirectoryJunction junction = new(directory.FullName);
			string alias = Path.Combine(junction.Path, "missing.pst");

			Assert.That(StorePath.AreEqual(path, alias), Is.False);
			Assert.That(File.Exists(alias), Is.False);
			Assert.That(File.Exists(path), Is.EqualTo(existingFile));
		}
		finally
		{
			directory.Delete(true);
		}
	}
}
