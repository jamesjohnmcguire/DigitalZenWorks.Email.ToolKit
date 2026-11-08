/////////////////////////////////////////////////////////////////////////////
// <copyright file="PublicApiCompatibilityTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Reflection;
using System.Text;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Guards released public signatures used by existing NuGet consumers.
/// </summary>
[Category("Compatibility")]
internal sealed class PublicApiCompatibilityTests
{
	/// <summary>
	/// Verifies that both migration overloads retain a void return type.
	/// </summary>
	/// <param name="includeEncoding">
	/// Whether to check the overload with an encoding parameter.
	/// </param>
	[TestCase(false)]
	[TestCase(true)]
	public void DbxDirectoryToPstRetainsVoidReturn(bool includeEncoding)
	{
		Type[] parameterTypes;

		if (includeEncoding)
		{
			parameterTypes =
				[typeof(string), typeof(string), typeof(Encoding)];
		}
		else
		{
			parameterTypes = [typeof(string), typeof(string)];
		}

		MethodInfo? method = typeof(Migrate).GetMethod(
			nameof(Migrate.DbxDirectoryToPst), parameterTypes);

		Assert.That(method, Is.Not.Null);
		Assert.That(method!.ReturnType, Is.EqualTo(typeof(void)));
	}

	/// <summary>
	/// Verifies the released file/path migration overload remains available
	/// without requiring the newer close-store argument.
	/// </summary>
	[Test]
	public void EmlFileToPstRetainsReleasedPathOverload()
	{
		MethodInfo? method = typeof(Migrate).GetMethod(
			nameof(Migrate.EmlFileToPst),
			BindingFlags.Public | BindingFlags.Static,
			[typeof(string), typeof(string)]);

		Assert.That(
			method,
			Is.Not.Null,
			"Restore the released EmlFileToPst(string, string) overload.");
		Assert.That(
			method!.ReturnType.FullName,
			Is.EqualTo(typeof(Outlook.MailItem).FullName));
	}

	/// <summary>
	/// Verifies the released migration overload accepting the adjust option
	/// remains available without requiring a close-store argument.
	/// </summary>
	[Test]
	public void EmlToPstRetainsReleasedAdjustOverload()
	{
		MethodInfo? method = typeof(Migrate).GetMethod(
			nameof(Migrate.EmlToPst),
			BindingFlags.Public | BindingFlags.Static,
			[typeof(string), typeof(string), typeof(bool)]);

		Assert.That(
			method,
			Is.Not.Null,
			"Restore the released EmlToPst(string, string, bool) overload.");
		Assert.That(method!.ReturnType, Is.EqualTo(typeof(bool)));
	}

	/// <summary>
	/// Verifies the released two-argument folder overload remains supported.
	/// An optional third argument does not retain its binary signature.
	/// </summary>
	[Test]
	public void OutlookFolderAddFolderRetainsReleasedTwoParameterOverload()
	{
		MethodInfo? method = typeof(OutlookFolder).GetMethod(
			nameof(OutlookFolder.AddFolder),
			BindingFlags.Public | BindingFlags.Static,
			[typeof(Outlook.MAPIFolder), typeof(string)]);

		Assert.That(
			method,
			Is.Not.Null,
			"Restore the released AddFolder(MAPIFolder, string) overload.");
		Assert.That(
			method!.ReturnType.FullName,
			Is.EqualTo(typeof(Outlook.MAPIFolder).FullName));
		Assert.That(
			method.IsDefined(typeof(ObsoleteAttribute)),
			Is.False,
			"Keep the two-argument convenience overload supported.");
	}

	/// <summary>
	/// Verifies new folder creation paths do not add public constructor
	/// overloads that make released null-argument calls ambiguous.
	/// </summary>
	[Test]
	public void OutlookFolderRetainsReleasedPublicConstructors()
	{
		Type type = typeof(OutlookFolder);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(type.GetConstructors(), Has.Length.EqualTo(2));
			Assert.That(
				type.GetConstructor([typeof(OutlookAccount)]), Is.Not.Null);
			Assert.That(
				type.GetConstructor(
					[typeof(OutlookAccount), typeof(Outlook.MAPIFolder)]),
				Is.Not.Null);
		}
	}

	/// <summary>
	/// Verifies that the legacy folder lookup method remains available.
	/// </summary>
	[Test]
	public void OutlookStoreRetainsLegacyGetFolderFromID()
	{
		MethodInfo? method = typeof(OutlookStore).GetMethod(
			"GetFolderFromID", [typeof(string), typeof(Outlook.Store)]);

		Assert.That(method, Is.Not.Null);
		Assert.That(
			method!.ReturnType.FullName,
			Is.EqualTo(typeof(Outlook.MAPIFolder).FullName));
	}

	/// <summary>
	/// Verifies new store creation paths preserve the released public
	/// constructor set and its null-argument overload resolution.
	/// </summary>
	[Test]
	public void OutlookStoreRetainsReleasedPublicConstructors()
	{
		Type type = typeof(OutlookStore);

		using (Assert.EnterMultipleScope())
		{
			Assert.That(type.GetConstructors(), Has.Length.EqualTo(1));
			Assert.That(
				type.GetConstructor([typeof(OutlookAccount)]), Is.Not.Null);
		}
	}
}
