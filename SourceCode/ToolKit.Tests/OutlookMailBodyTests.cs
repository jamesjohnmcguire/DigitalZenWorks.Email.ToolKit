/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookMailBodyTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.IO;
using NUnit.Framework;

/// <summary>
/// Tests OutlookMail's file-writing behavior without activating Outlook.
/// </summary>
internal sealed class OutlookMailBodyTests
{
	private DirectoryInfo directory = null!;

	/// <summary>
	/// Creates an isolated export directory.
	/// </summary>
	[SetUp]
	public void SetUp()
	{
		directory = Directory.CreateTempSubdirectory("Email.ToolKit.Bodies.");
	}

	/// <summary>
	/// Removes this test's files.
	/// </summary>
	[TearDown]
	public void TearDown()
	{
		directory.Delete(true);
	}

	/// <summary>
	/// Verifies an existing file prevents replacement or a partial new pair.
	/// </summary>
	/// <param name="extension">The existing output format.</param>
	[TestCase(".txt")]
	[TestCase(".html")]
	public void WriteDoesNotOverwriteExistingExport(string extension)
	{
		string existing = Path.Combine(directory.FullName, "entry" + extension);
		File.WriteAllText(existing, "Previous export.");
		bool written = WriteBodies("text", "<p>HTML</p>");

		Assert.That(written, Is.False);
		Assert.That(File.ReadAllText(existing), Is.EqualTo("Previous export."));
		string[] files = Directory.GetFiles(directory.FullName);
		Assert.That(files, Has.Length.EqualTo(1));
	}

	/// <summary>
	/// Verifies sender metadata cannot terminate the HTML comment early.
	/// </summary>
	[Test]
	public void WriteEscapesHtmlMetadata()
	{
		WriteBodies(
			"text", "<p>HTML</p>", sender: "Recruiter -->\r\nInjected");

		string path = Path.Combine(directory.FullName, "entry.html");
		string html = File.ReadAllText(path);
		string[] lines = html.Split('\n');
		Assert.That(lines[0], Does.StartWith("<!-- Sent: "));

		// Flatten sender CR/LF so the writer's closing delimiter remains on
		// the second metadata line, after the complete sender value.
		Assert.That(lines[1], Does.EndWith(" -->"));

		// The sender's own --> is data: escape it so it cannot close the
		// comment early. Only the writer appends the actual --> delimiter.
		Assert.That(lines[1], Does.Contain("&#45;&#45;&gt;"));
		Assert.That(lines[2], Is.Empty);
		Assert.That(html, Does.EndWith("<p>HTML</p>"));
	}

	/// <summary>
	/// Verifies absent HTML does not generate an HTML approximation.
	/// </summary>
	/// <param name="html">The absent HTML value.</param>
	[TestCase((string?)null)]
	[TestCase("")]
	public void WriteOmitsUnavailableHtml(string? html)
	{
		WriteBodies(null, html);

		string textPath = Path.Combine(directory.FullName, "entry.txt");
		string htmlPath = Path.Combine(directory.FullName, "entry.html");
		Assert.That(File.Exists(textPath), Is.True);
		Assert.That(File.Exists(htmlPath), Is.False);
	}

	/// <summary>
	/// Verifies UTF-8 output preserves the exact body after metadata.
	/// </summary>
	[Test]
	public void WritePreservesBodiesAndMetadata()
	{
		const string text = "Café\r\nDeveloper position\n";
		const string html = "<html><body>Café &amp; jobs</body></html>";
		bool written = WriteBodies(text, html);

		string textPath = Path.Combine(directory.FullName, "entry.txt");
		string htmlPath = Path.Combine(directory.FullName, "entry.html");
		string textOutput = File.ReadAllText(textPath);
		string htmlOutput = File.ReadAllText(htmlPath);
		const string metadata = "Sent: 2026-11-18T10:20:30.0000000Z\n" +
			"Sender: Recruiter <jobs@example.com>\n\n";
		Assert.That(written, Is.True);
		Assert.That(textOutput, Is.EqualTo(metadata + text));
		Assert.That(htmlOutput, Does.StartWith(
			"<!-- Sent: 2026-11-18T10:20:30.0000000Z\n" +
			"Sender: Recruiter &lt;jobs@example.com&gt; -->\n\n"));
		Assert.That(htmlOutput, Does.EndWith(html));
		byte[] bytes = File.ReadAllBytes(textPath);
		Assert.That(bytes[0], Is.EqualTo((byte)'S'));
	}

	private bool WriteBodies(
		string? text,
		string? html,
		string sender = "Recruiter <jobs@example.com>")
	{
		bool written = OutlookMail.WriteBodiesWrapper(
			"entry",
			"2026-11-18T10:20:30.0000000Z",
			sender,
			text!,
			html!,
			directory.FullName);

		return written;
	}
}
