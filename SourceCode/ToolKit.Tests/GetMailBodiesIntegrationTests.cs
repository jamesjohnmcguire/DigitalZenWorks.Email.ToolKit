/////////////////////////////////////////////////////////////////////////////
// <copyright file="GetMailBodiesIntegrationTests.cs"
// company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using DigitalZenWorks.CommandLine.Commands;
using DigitalZenWorks.Email.ToolKit.Application;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;
using Program = DigitalZenWorks.Email.ToolKit.Application.Program;

/// <summary>
/// Exercises body export against real mail in isolated test-owned PSTs.
/// </summary>
[Apartment(ApartmentState.STA)]
[Category("Integration")]
[NonParallelizable]
internal sealed class GetMailBodiesIntegrationTests
{
	private DirectoryInfo directory = null!;

	/// <summary>
	/// Verifies the command opens a PST and leaves it attached.
	/// </summary>
	[Test]
	public void ExportAttachesSourceStore()
	{
		using DuplicateTestData data = new();
		Outlook.MailItem mail = data.CreateMail();
		DuplicateTestData.Release(mail);
		Assert.That(data.Session.RemoveStore(data.StorePath), Is.True);
		string output = Path.Combine(directory.FullName, "output");
		string[] parameters = [data.StorePath, "Candidates", output];
		Command command = new("get-mail-bodies", [], parameters);
		OutlookService service = CreateConnectedService(data);

		int result = Commands.GetMailBodies(service, command);

		Assert.That(result, Is.Zero);
		Assert.That(Directory.GetFiles(output), Has.Length.EqualTo(2));
		Outlook.Store? attached = data.Session.GetStore(data.StorePath, false);
		if (attached != null)
		{
			DuplicateTestData.Release(attached);
		}

		Assert.That(attached, Is.Not.Null);
	}

	/// <summary>
	/// Verifies a missing folder throws without creating source folders.
	/// </summary>
	[Test]
	public void ExportDoesNotCreateMissingFolder()
	{
		using DuplicateTestData data = new();
		Outlook.MAPIFolder folder = data.OpenFolder();
		Outlook.Folders children = folder.Folders;
		try
		{
			int originalCount = children.Count;
			string output = Path.Combine(directory.FullName, "output");
			string[] parameters =
				[data.StorePath, "Candidates/Missing", output];
			Command command = new("get-mail-bodies", [], parameters);
			OutlookService service = CreateConnectedService(data);

			Assert.Throws<COMException>(() =>
				Commands.GetMailBodies(service, command));
			Assert.That(children.Count, Is.EqualTo(originalCount));
			Assert.That(Directory.Exists(output), Is.False);
			Assert.That(service.IsConnected, Is.True);
		}
		finally
		{
			DuplicateTestData.Release(children, folder);
		}
	}

	/// <summary>
	/// Verifies the real command-line process calls the library exporter.
	/// </summary>
	[Test]
	public void ExportRunsFromCommandLineProcess()
	{
		using DuplicateTestData data = new();
		Outlook.MailItem mail = data.CreateMail();
		string entryId = mail.EntryID;
		DuplicateTestData.Release(mail);
		string output = Path.Combine(directory.FullName, "output");

		ProcessStartInfo startInfo = new("dotnet");
		startInfo.UseShellExecute = false;
		startInfo.CreateNoWindow = true;
		startInfo.ArgumentList.Add(typeof(Program).Assembly.Location);
		startInfo.ArgumentList.Add("get-mail-bodies");
		startInfo.ArgumentList.Add(data.StorePath);
		startInfo.ArgumentList.Add("Candidates");
		startInfo.ArgumentList.Add(output);
		using Process process = Process.Start(startInfo)!;
		bool exited = process.WaitForExit(60000);
		if (exited == false)
		{
			// Terminate only this test's CLI host, never Outlook.
			process.Kill();
			process.WaitForExit();
			Assert.Fail("The body-export CLI exceeded its test watchdog.");
		}

		Assert.That(process.ExitCode, Is.Zero);
		string textPath = Path.Combine(output, entryId + ".txt");
		string htmlPath = Path.Combine(output, entryId + ".html");
		Assert.That(File.Exists(textPath), Is.True);
		Assert.That(File.Exists(htmlPath), Is.True);
	}

	/// <summary>
	/// Verifies non-mail items are skipped without failing the export.
	/// </summary>
	[Test]
	public void ExportSkipsNonMailItems()
	{
		using DuplicateTestData data = new();
		Outlook.MAPIFolder folder = data.OpenFolder();
		Outlook.Items items = folder.Items;
		Outlook.PostItem post =
			(Outlook.PostItem)items.Add(Outlook.OlItemType.olPostItem);
		try
		{
			post.Subject = "Non-mail fixture";
			post.Save();
			string output = Path.Combine(directory.FullName, "output");
			OutlookStore store = OutlookStore.FromSession(data.Session);

			bool result = store.GetMailBodies(
				data.StorePath, "Candidates", output);

			Assert.That(result, Is.True);
			Assert.That(Directory.GetFiles(output), Is.Empty);
		}
		finally
		{
			DuplicateTestData.Release(post, items, folder);
		}
	}

	/// <summary>
	/// Verifies body export resolves identical root/child names and preserves
	/// the source mail and an already-attached Outlook store.
	/// </summary>
	[Test]
	public void ExportWritesBodiesAndPreservesSource()
	{
		using DuplicateTestData data = new();
		SetEmploymentFolderNames(data);
		Outlook.MailItem mail = data.CreateMail();
		try
		{
			mail.UnRead = true;
			mail.Save();
			string entryId = mail.EntryID;
			string text = mail.Body;
			string html = mail.HTMLBody;
			string output = Path.Combine(directory.FullName, "output");
			string[] parameters =
				[data.StorePath, "Employment/Employment", output];

			Command command = new("get-mail-bodies", [], parameters);
			OutlookService service = CreateConnectedService(data);

			int result = Commands.GetMailBodies(service, command);

			Assert.That(result, Is.Zero);
			string textPath = Path.Combine(output, entryId + ".txt");
			string htmlPath = Path.Combine(output, entryId + ".html");
			string textOutput = File.ReadAllText(textPath);
			string htmlOutput = File.ReadAllText(htmlPath);
			Assert.That(textOutput, Does.StartWith("Sent: "));
			Assert.That(textOutput, Does.Contain("Safety Sender"));
			Assert.That(textOutput, Does.EndWith(text));
			Assert.That(htmlOutput, Does.StartWith("<!-- Sent: "));
			Assert.That(htmlOutput, Does.EndWith(html));
			Assert.That(Directory.GetFiles(output), Has.Length.EqualTo(2));
			Assert.That(mail.Body, Is.EqualTo(text));
			Assert.That(mail.HTMLBody, Is.EqualTo(html));
			Assert.That(mail.UnRead, Is.True);
			Assert.That(mail.EntryID, Is.EqualTo(entryId));
			Assert.That(service.IsConnected, Is.True);
			Outlook.Store? store = data.Session.GetStore(data.StorePath, false);
			try
			{
				Assert.That(store, Is.Not.Null);
			}
			finally
			{
				if (store != null)
				{
					DuplicateTestData.Release(store);
				}
			}
		}
		finally
		{
			DuplicateTestData.Release(mail);
		}
	}

	/// <summary>
	/// Creates an isolated destination for exported files.
	/// </summary>
	[SetUp]
	public void SetUp()
	{
		directory = Directory.CreateTempSubdirectory("Email.ToolKit.Export.");
	}

	/// <summary>
	/// Removes the exported files after COM cleanup.
	/// </summary>
	[TearDown]
	public void TearDown()
	{
		directory.Delete(true);
	}

	private static OutlookService CreateConnectedService(
		DuplicateTestData data)
	{
		OutlookConnection connection = new(data.Session);
		FakeOutlookFactory factory = new();
		factory.IsAvailable = true;
		factory.Connection = connection;
		OutlookService service = new();
		bool connected = service.Connect(factory);
		Assert.That(connected, Is.True);
		return service;
	}

	private static void SetEmploymentFolderNames(DuplicateTestData data)
	{
		Outlook.Store store = data.Session.GetStore(data.StorePath, false)!;
		Outlook.MAPIFolder root = store.GetRootFolder();
		Outlook.MAPIFolder folder = data.OpenFolder();
		try
		{
			root.Name = "Employment";
			folder.Name = "Employment";
		}
		finally
		{
			DuplicateTestData.Release(folder, root, store);
		}
	}
}
