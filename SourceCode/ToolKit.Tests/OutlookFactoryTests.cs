/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFactoryTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.IO;
using System.Threading;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Exercises the real factory against the available Outlook installation.
/// </summary>
[Category("Integration")]
[Apartment(ApartmentState.STA)]
[NonParallelizable]
internal sealed class OutlookFactoryTests
{
	/// <summary>
	/// Verifies the availability probe succeeds with Outlook available.
	/// </summary>
	[Test]
	public void CanCreateApplicationWhenOutlookIsAvailable()
	{
		using OutlookTestContext context = new();
		OutlookFactory factory = new();

		Assert.That(factory.CanCreateApplication(33), Is.True);
		Outlook.Stores stores = context.Track(context.NameSpace.Stores);
		Assert.That(stores.Count, Is.GreaterThan(0));
	}

	/// <summary>
	/// Verifies factory-created connections expose stable, usable sessions.
	/// </summary>
	[Test]
	public void CreateConnectionExposesStableUsableSession()
	{
		using OutlookTestContext context = new();
		OutlookFactory factory = new();
		IOutlookConnection? connection = factory.CreateConnection();

		using IDisposable? cleanup = connection as IDisposable;

		Assert.That(connection, Is.Not.Null);
		IOutlookSession? first = connection!.Session;
		Assert.That(first, Is.Not.Null);
		Assert.That(connection.Session, Is.SameAs(first));

		Outlook.MailItem original = context.Track(
			(Outlook.MailItem)context.Application.CreateItem(
				Outlook.OlItemType.olMailItem));
		original.Subject = "Factory session test";
		string path = Path.Combine(context.DirectoryPath, "factory.msg");
		original.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);

		object? result = first!.OpenSharedItem(path);

		Assert.That(result, Is.InstanceOf<Outlook.MailItem>());
		Outlook.MailItem opened = context.Track((Outlook.MailItem)result!);
		Assert.That(opened.Subject, Is.EqualTo(original.Subject));
		opened.Close(Outlook.OlInspectorClose.olDiscard);
		original.Close(Outlook.OlInspectorClose.olDiscard);
	}
}
