/////////////////////////////////////////////////////////////////////////////
// <copyright file="DuplicateHashTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using NUnit.Framework;

/// <summary>
/// Checks duplicate grouping when comparison data is unavailable.
/// </summary>
internal sealed class DuplicateHashTests
{
	private static readonly string?[] MissingHashes = [null, string.Empty];

	/// <summary>
	/// An unavailable hash must never join a valid duplicate group.
	/// </summary>
	/// <param name="hash">The failed or empty hash.</param>
	[TestCaseSource(nameof(MissingHashes))]
	public void MissingHashDoesNotJoinDuplicateGroup(string? hash)
	{
		Dictionary<string, IList<string>> groups = new()
		{
			["valid-hash"] = new List<string> { "keeper" }
		};
		MethodInfo? method = typeof(OutlookFolder).GetMethod(
			"AddHashToTable", BindingFlags.NonPublic | BindingFlags.Static);
		Assert.That(method, Is.Not.Null);
		object? result = method!.Invoke(null, [groups, hash, "unreadable"]);
		Assert.That(result, Is.SameAs(groups));
		Assert.That(groups, Has.Count.EqualTo(1));
		Assert.That(groups["valid-hash"], Has.Count.EqualTo(1));
		Assert.That(groups["valid-hash"][0], Is.EqualTo("keeper"));
	}

	/// <summary>
	/// A missing item must not produce a shared placeholder hash.
	/// </summary>
	/// <returns>The asynchronous test.</returns>
	[Test]
	public async Task NullItemsDoNotProduceUsableHashes()
	{
		OutlookItem item = new(null!);
		Assert.That(item.Hash, Is.Null);
		string? hash = await OutlookItem.GetHashAsync(null!)
			.ConfigureAwait(false);
		Assert.That(hash, Is.Null);
	}
}
