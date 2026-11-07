/////////////////////////////////////////////////////////////////////////////
// <copyright file="LegacyMigrationTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.IO;
using System.Threading;
using DigitalZenWorks.Common.Utilities;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Checks restored overload behavior against real Outlook and temporary PSTs.
/// </summary>
[Apartment(ApartmentState.STA)]
[Category("Integration")]
[NonParallelizable]
internal sealed class LegacyMigrationTests
{
	/// <summary>
	/// Verifies the convenience overload retains a duplicate-looking name.
	/// </summary>
	[Test]
	public void AddFolderPreservesNameWithoutNormalization()
	{
		using OutlookTestContext context = new();
		string path = context.StorePath("folder-name.pst");
		OutlookSession session = new(context.NameSpace);
		Outlook.Store store = session.GetStore(path)!;
		context.Track(store);
		Outlook.MAPIFolder root = store.GetRootFolder();
		context.Track(root);

		Outlook.MAPIFolder folder =
			OutlookFolder.AddFolder(root, "Archive (1)");
		context.TrackIfNotNull(folder);

		Assert.That(folder, Is.Not.Null);
		Assert.That(folder.Name, Is.EqualTo("Archive (1)"));
	}

	/// <summary>
	/// Verifies the released file overload imports without detaching the PST.
	/// </summary>
	[Test]
	public void EmlFileToPstLeavesStoreAttached()
	{
		using OutlookTestContext context = new();
		string path = context.StorePath("legacy-file.pst");
		string source = Path.Combine(context.DirectoryPath, "source.eml");
		bool extracted = FileUtils.CreateFileFromEmbeddedResource(
			"ToolKit.Tests.TestEmail.eml", source);
		Assert.That(extracted, Is.True);

#pragma warning disable CS0618 // Exercise the released compatibility API.
		Outlook.MailItem item = Migrate.EmlFileToPst(source, path);
#pragma warning restore CS0618
		context.TrackIfNotNull(item);

		Assert.That(item, Is.Not.Null);
		Assert.That(context.CountStores(path), Is.EqualTo(1));
	}

	/// <summary>
	/// Verifies adjustment is forwarded and the destination stays attached.
	/// </summary>
	/// <param name="adjust">Whether to skip the interim source folder.</param>
	[TestCase(false)]
	[TestCase(true)]
	public void EmlToPstPreservesAdjustmentAndLeavesStoreAttached(bool adjust)
	{
		using OutlookTestContext context = new();
		string path = context.StorePath("legacy-directory.pst");
		string source = Path.Combine(context.DirectoryPath, "Local Folders");
		Directory.CreateDirectory(source);
		string emlPath = Path.Combine(source, "source.eml");
		bool extracted = FileUtils.CreateFileFromEmbeddedResource(
			"ToolKit.Tests.TestEmail.eml", emlPath);
		Assert.That(extracted, Is.True);

#pragma warning disable CS0618 // Exercise the released compatibility API.
		bool result = Migrate.EmlToPst(source, path, adjust);
#pragma warning restore CS0618

		Assert.That(result, Is.True);
		Assert.That(context.CountStores(path), Is.EqualTo(1));
		OutlookSession session = new(context.NameSpace);
		Outlook.Store store = session.GetStore(path)!;
		context.Track(store);
		Outlook.MAPIFolder root = store.GetRootFolder();
		context.Track(root);
		Outlook.MAPIFolder destination = root;

		if (adjust == false)
		{
			Outlook.Folders folders = root.Folders;
			context.Track(folders);
			destination = folders["Local Folders"];
			context.Track(destination);
		}

		Outlook.Items items = destination.Items;
		context.Track(items);
		Assert.That(items.Count, Is.EqualTo(1));
	}
}
