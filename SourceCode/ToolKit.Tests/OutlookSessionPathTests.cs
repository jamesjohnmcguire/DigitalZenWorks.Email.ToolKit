/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookSessionPathTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.IO;
using System.Reflection;
using NUnit.Framework;

/// <summary>
/// Tests PST path selection using real temporary files without Outlook.
/// </summary>
internal sealed class OutlookSessionPathTests
{
	/// <summary>
	/// Verifies lookup preserves an existing extensionless file when no PST
	/// alias exists, while creation and existing aliases use the PST suffix.
	/// </summary>
	/// <param name="originalExists">Whether the supplied file exists.</param>
	/// <param name="pstExists">Whether the suffixed PST file exists.</param>
	/// <param name="createIfMissing">Whether store creation is requested.</param>
	/// <returns>The selected file name.</returns>
	[TestCase(false, false, false, ExpectedResult = "archive.pst")]
	[TestCase(false, false, true, ExpectedResult = "archive.pst")]
	[TestCase(false, true, false, ExpectedResult = "archive.pst")]
	[TestCase(false, true, true, ExpectedResult = "archive.pst")]
	[TestCase(true, false, false, ExpectedResult = "archive")]
	[TestCase(true, false, true, ExpectedResult = "archive.pst")]
	[TestCase(true, true, false, ExpectedResult = "archive.pst")]
	[TestCase(true, true, true, ExpectedResult = "archive.pst")]
	public string GetAdjustedPstPathSelectsExpectedFile(
		bool originalExists,
		bool pstExists,
		bool createIfMissing)
	{
		DirectoryInfo directory =
			Directory.CreateTempSubdirectory("Email.ToolKit.Path.");

		try
		{
			string path = Path.Combine(directory.FullName, "archive");
			string pstPath = path + ".pst";

			if (originalExists)
			{
				File.WriteAllText(path, "Existing file.");
			}

			if (pstExists)
			{
				File.WriteAllText(pstPath, "Existing PST alias.");
			}

			// Exercise the production path decision without widening its
			// visibility or creating fake Outlook COM objects.
			MethodInfo? method = typeof(OutlookSession).GetMethod(
				"GetAdjustedPstPath",
				BindingFlags.NonPublic | BindingFlags.Static);
			Assert.That(method, Is.Not.Null);

			object? actual = method!.Invoke(null, [path, createIfMissing]);
			Assert.That(actual, Is.TypeOf<string>());
			string adjusted = (string)actual!;

			using (Assert.EnterMultipleScope())
			{
				Assert.That(
					Path.GetDirectoryName(adjusted),
					Is.EqualTo(directory.FullName));
				Assert.That(File.Exists(path), Is.EqualTo(originalExists));
				Assert.That(File.Exists(pstPath), Is.EqualTo(pstExists));
			}

			return Path.GetFileName(adjusted);
		}
		finally
		{
			directory.Delete(true);
		}
	}
}
