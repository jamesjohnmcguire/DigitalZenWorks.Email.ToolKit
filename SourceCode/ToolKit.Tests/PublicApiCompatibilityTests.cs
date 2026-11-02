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
}
