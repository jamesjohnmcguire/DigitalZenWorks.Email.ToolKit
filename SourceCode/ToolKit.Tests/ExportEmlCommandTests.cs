/////////////////////////////////////////////////////////////////////////////
// <copyright file="ExportEmlCommandTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.IO;
using System.Threading.Tasks;
using NUnit.Framework;
using Program = DigitalZenWorks.Email.ToolKit.Application.Program;

/// <summary>
/// Tests EML export command parsing without starting Outlook.
/// </summary>
internal sealed class ExportEmlCommandTests
{
	private DirectoryInfo directory = null!;
	private string outputPath = string.Empty;
	private string storePath = string.Empty;

	/// <summary>
	/// Verifies separate source arguments forward either timeout option once.
	/// </summary>
	/// <param name="option">The timeout option spelling.</param>
	/// <returns>The asynchronous test operation.</returns>
	[TestCase("-t")]
	[TestCase("--timeout")]
	public async Task ExportEmlAcceptsSeparateArguments(string option)
	{
		FakeOutlookService outlook = new();
		string[] parameters =
			[storePath, "Employment/Employment", outputPath];

		string[] arguments =
			["export-eml", option, "37", .. parameters];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(outlook.ConnectCallCount, Is.EqualTo(1));
		Assert.That(outlook.LastTimeoutSeconds, Is.EqualTo(37));
		Assert.That(outlook.SessionAccessCount, Is.Zero);
		Assert.That(result, Is.Not.Zero);
		Assert.That(Directory.Exists(outputPath), Is.False);
	}

	/// <summary>
	/// Verifies an invalid timeout never attempts a connection.
	/// </summary>
	/// <param name="value">The invalid timeout value.</param>
	/// <returns>The asynchronous test operation.</returns>
	[TestCase("-1")]
	[TestCase("invalid")]
	[TestCase("2147484")]
	public async Task ExportEmlRejectsInvalidTimeout(string value)
	{
		FakeOutlookService outlook = new();
		string[] arguments =
		[
			"export-eml", storePath, "Employment", outputPath,
			"--timeout", value
		];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
	}

	/// <summary>
	/// Verifies two arguments cannot replace the three required arguments.
	/// </summary>
	/// <param name="source">The incomplete source argument.</param>
	/// <returns>The asynchronous test operation.</returns>
	[TestCase("::Employment")]
	[TestCase("archive.pst::")]
	[TestCase("archive.pst::Employment")]
	[TestCase("archive.pst")]
	public async Task ExportEmlRejectsMalformedSource(string source)
	{
		FakeOutlookService outlook = new();
		string[] arguments = ["export-eml", source, outputPath];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
	}

	/// <summary>
	/// Verifies exporting cannot create a missing source PST.
	/// </summary>
	/// <returns>The asynchronous test operation.</returns>
	[Test]
	public async Task ExportEmlRejectsMissingStore()
	{
		File.Delete(storePath);
		FakeOutlookService outlook = new();
		string[] arguments =
			["export-eml", storePath, "Employment", outputPath];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
		Assert.That(File.Exists(storePath), Is.False);
	}

	/// <summary>
	/// Creates ordinary files for syntax checks that never reach Outlook.
	/// </summary>
	[SetUp]
	public void SetUp()
	{
		directory = Directory.CreateTempSubdirectory("Email.ToolKit.Export.");
		storePath = Path.Combine(directory.FullName, "source.pst");
		outputPath = Path.Combine(directory.FullName, "output");
		File.WriteAllText(storePath, "Not opened by the failing service.");
	}

	/// <summary>
	/// Deletes this fixture's temporary files.
	/// </summary>
	[TearDown]
	public void TearDown()
	{
		directory.Delete(true);
	}
}
