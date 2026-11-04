/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandLineTimeoutTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Globalization;
using System.Threading.Tasks;
using NUnit.Framework;
using Program = DigitalZenWorks.Email.ToolKit.Application.Program;

/// <summary>
/// Tests CLI timeout parsing and forwarding without activating Outlook.
/// </summary>
internal sealed class CommandLineTimeoutTests
{
	private static readonly TestCaseData[] OutlookCommands =
	[
		new("add-pst", new[] { "archive.pst" }),
		new("dbx-to-pst", new[] { "Inbox.dbx" }),
		new("details", new[] { "archive.pst", "entry-id" }),
		new("eml-to-pst", new[] { "message.eml", "archive.pst" }),
		new("list-folders", new[] { "archive.pst" }),
		new("list-ids", new[] { "archive.pst" }),
		new("list-top-senders", new[] { "archive.pst" }),
		new("list-total-duplicates", new[] { "archive.pst" }),
		new("merge-folders", Array.Empty<string>()),
		new("merge-stores", new[] { "source.pst", "destination.pst" }),
		new(
			"move-folder",
			new[] { "source.pst", "Inbox", "destination.pst", "Archive" }),
		new("remove-duplicates", Array.Empty<string>()),
		new("remove-empty-folders", new[] { "archive.pst" })
	];

	/// <summary>
	/// Verifies help is available without attempting an Outlook connection.
	/// </summary>
	/// <param name="argument">The help command or one of its aliases.</param>
	/// <returns>A task representing the asynchronous test.</returns>
	[TestCase("help")]
	[TestCase("--help")]
	[TestCase("-h")]
	[TestCase("-?")]
	public async Task HelpDoesNotConnect(string argument)
	{
		FakeOutlookService outlook = new();
		string[] arguments = [argument];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies invalid commands and missing arguments cannot start Outlook.
	/// </summary>
	/// <param name="command">An unknown command or one requiring arguments.
	/// </param>
	/// <returns>A task representing the asynchronous test.</returns>
	[TestCase("add-pst")]
	[TestCase("merge-stores")]
	[TestCase("unknown-command")]
	public async Task InvalidArgumentsDoNotConnect(string command)
	{
		FakeOutlookService outlook = new();
		string[] arguments = [command];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies both option forms accept the supported timeout boundaries.
	/// Zero means no wait for the probe, not an infinite timeout.
	/// </summary>
	/// <param name="option">The short or long timeout option.</param>
	/// <param name="seconds">The requested probe timeout.</param>
	/// <returns>A task representing the asynchronous test.</returns>
	[Test]
	public async Task OutlookTimeoutAcceptsSupportedValues(
		[Values("--timeout", "-t")] string option,
		[Values(0, 1, 2147483)] int seconds)
	{
		FakeOutlookService outlook = new();
		string value = seconds.ToString(CultureInfo.InvariantCulture);
		string[] arguments = ["merge-folders", option, value];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(outlook.ConnectCallCount, Is.EqualTo(1));
		Assert.That(outlook.LastTimeoutSeconds, Is.EqualTo(seconds));
		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies a timeout value cannot satisfy a required positional argument.
	/// </summary>
	/// <param name="option">The short or long timeout option.</param>
	/// <returns>A task representing the asynchronous test.</returns>
	[TestCase("--timeout")]
	[TestCase("-t")]
	public async Task OutlookTimeoutDoesNotReplaceRequiredParameter(
		string option)
	{
		FakeOutlookService outlook = new();
		string[] arguments = ["add-pst", option, "37"];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies every Outlook command forwards the explicit timeout once and
	/// returns failure without requesting a session when connection fails.
	/// </summary>
	/// <param name="command">The Outlook command to execute.</param>
	/// <param name="parameters">Valid positional arguments for the command.
	/// </param>
	/// <returns>A task representing the asynchronous test.</returns>
	[TestCaseSource(nameof(OutlookCommands))]
	public async Task OutlookTimeoutIsForwardedByEveryCommand(
		string command, string[] parameters)
	{
		FakeOutlookService outlook = new();
		string[] arguments =
			[command, .. parameters, "--timeout", "37"];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(outlook.ConnectCallCount, Is.EqualTo(1));
		Assert.That(outlook.LastTimeoutSeconds, Is.EqualTo(37));
		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies invalid timeout values fail before contacting Outlook.
	/// </summary>
	/// <param name="option">The short or long timeout option.</param>
	/// <param name="value">An invalid timeout value.</param>
	/// <returns>A task representing the asynchronous test.</returns>
	[Test]
	public async Task OutlookTimeoutRejectsInvalidValues(
		[Values("--timeout", "-t")] string option,
		[Values("-1", "2147484", "2147483648", "abc", "1.5", "", " ")]
		string value)
	{
		FakeOutlookService outlook = new();
		string[] arguments = ["merge-folders", option, value];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies an option without its value fails before contacting Outlook.
	/// </summary>
	/// <param name="option">The short or long timeout option.</param>
	/// <returns>A task representing the asynchronous test.</returns>
	[TestCase("--timeout")]
	[TestCase("-t")]
	public async Task OutlookTimeoutRejectsMissingValue(string option)
	{
		FakeOutlookService outlook = new();
		string[] arguments = ["add-pst", "archive.pst", option];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies another option cannot be consumed as the timeout value.
	/// </summary>
	/// <param name="option">The short or long timeout option.</param>
	/// <returns>A task representing the asynchronous test.</returns>
	[TestCase("--timeout")]
	[TestCase("-t")]
	public async Task OutlookTimeoutRejectsOptionAsValue(string option)
	{
		FakeOutlookService outlook = new();
		string[] arguments = ["merge-folders", option, "--dryrun"];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.ConnectCallCount, Is.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}

	/// <summary>
	/// Verifies omitting the option preserves the ten-second probe timeout for
	/// every Outlook command through the shared connection step.
	/// </summary>
	/// <param name="command">The Outlook command to execute.</param>
	/// <param name="parameters">Valid positional arguments for the command.
	/// </param>
	/// <returns>A task representing the asynchronous test.</returns>
	[TestCaseSource(nameof(OutlookCommands))]
	public async Task OutlookTimeoutUsesDefaultForEveryCommand(
		string command, string[] parameters)
	{
		FakeOutlookService outlook = new();
		string[] arguments = [command, .. parameters];

		int result = await Program.RunAsync(arguments, outlook).
			ConfigureAwait(false);

		Assert.That(outlook.ConnectCallCount, Is.EqualTo(1));
		Assert.That(outlook.LastTimeoutSeconds, Is.EqualTo(10));
		Assert.That(result, Is.Not.Zero);
		Assert.That(outlook.SessionAccessCount, Is.Zero);
	}
}
