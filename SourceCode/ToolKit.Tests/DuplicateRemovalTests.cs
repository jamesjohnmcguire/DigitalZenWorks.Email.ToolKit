/////////////////////////////////////////////////////////////////////////////
// <copyright file="DuplicateRemovalTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;
using Program = DigitalZenWorks.Email.ToolKit.Application.Program;

/// <summary>
/// Checks real survivors and deletion guards in isolated PSTs.
/// </summary>
[Apartment(ApartmentState.STA)]
[Category("Integration")]
[NonParallelizable]
internal sealed class DuplicateRemovalTests
{
	/// <summary>The CLI must ignore flush during a dry run.</summary>
	/// <param name="flush">Whether flush was also requested.</param>
	/// <returns>The asynchronous test.</returns>
	[TestCase(false)]
	[TestCase(true)]
	public async Task CliDryRunPreservesItemsAndDeletedSentinel(bool flush)
	{
		using DuplicateTestData data = new();
		CreateMixedSet(data);
		Outlook.MailItem sentinel = data.CreateDeletedSentinel();
		DuplicateTestData.Release(sentinel);
		Assert.That(data.GetDeletedCount(), Is.EqualTo(1));
		HashSet<string> before = data.GetItemIds();
		List<string> arguments = ["remove-duplicates", data.StorePath];
		arguments.Add("--dryrun");

		if (flush)
		{
			arguments.Add("--flush");
		}

		OutlookService service = new();

		try
		{
			int result = await Program.RunAsync(arguments.ToArray(), service)
				.ConfigureAwait(true);
			Assert.That(result, Is.Zero);
			HashSet<string> after = data.GetItemIds();
			Assert.That(after, Is.EquivalentTo(before));
			Assert.That(data.GetDeletedCount(), Is.EqualTo(1));
		}
		finally
		{
			service.Disconnect();
		}
	}

	/// <summary>Both public deletion overloads must honor dry run.</summary>
	/// <param name="useStore">
	/// Whether to call the store-wrapper overload.
	/// </param>
	/// <param name="dryRun">Whether deletion is disabled.</param>
	[TestCase(false, false)]
	[TestCase(false, true)]
	[TestCase(true, false)]
	[TestCase(true, true)]
	public void DeleteDuplicateHonorsDryRun(bool useStore, bool dryRun)
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = DuplicateTestData.Copy(first);
		string keeperId = first.EntryID;
		string duplicateId = second.EntryID;
		string synopsis = OutlookMail.GetSynopses(first);

		try
		{
			string firstHash = DuplicateTestData.RequiredHash(first);
			string secondHash = DuplicateTestData.RequiredHash(second);
			Assert.That(secondHash, Is.EqualTo(firstHash));
			Assert.That(synopsis, Is.Not.Null.And.Not.Empty);
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}

		DeleteCandidate(data, duplicateId, synopsis, useStore, dryRun);
		HashSet<string> survivors = data.GetItemIds();
		Assert.That(survivors, Does.Contain(keeperId));
		Assert.That(survivors.Contains(duplicateId), Is.EqualTo(dryRun));
		Assert.That(data.GetDeletedCount(), Is.EqualTo(dryRun ? 0 : 1));
	}

	/// <summary>
	/// A synopsis mismatch must prevent deletion even if requested.
	/// </summary>
	/// <param name="useStore">
	/// Whether to call the store-wrapper overload.
	/// </param>
	[TestCase(false)]
	[TestCase(true)]
	public void DeleteDuplicateRejectsDifferentSynopsis(bool useStore)
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = DuplicateTestData.Copy(first);
		string duplicateId = second.EntryID;
		string keeperSynopsis = OutlookMail.GetSynopses(first);

		try
		{
			second.Subject = "A different item must survive";
			second.Save();
			string candidateSynopsis = OutlookMail.GetSynopses(second);
			Assert.That(candidateSynopsis, Is.Not.EqualTo(keeperSynopsis));
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}

		HashSet<string> before = data.GetItemIds();
		DeleteCandidate(data, duplicateId, keeperSynopsis, useStore, false);
		HashSet<string> after = data.GetItemIds();
		Assert.That(after, Is.EquivalentTo(before));
		Assert.That(data.GetDeletedCount(), Is.Zero);
	}

	/// <summary>A dry run must preserve every original item.</summary>
	/// <param name="useAsync">Whether to use asynchronous removal.</param>
	/// <returns>The asynchronous test.</returns>
	[TestCase(false)]
	[TestCase(true)]
	public async Task RemoveDuplicatesHonorsDryRun(bool useAsync)
	{
		using DuplicateTestData data = new();
		CreateMixedSet(data);
		HashSet<string> before = data.GetItemIds();
		Assert.That(before, Has.Count.EqualTo(3));

		await data.RemoveDuplicatesAsync(useAsync, true)
			.ConfigureAwait(true);

		HashSet<string> after = data.GetItemIds();
		Assert.That(after, Is.EquivalentTo(before));
		Assert.That(data.GetDeletedCount(), Is.Zero);
	}

	/// <summary>
	/// Only one real duplicate may disappear; the unique item stays.
	/// </summary>
	/// <param name="useAsync">Whether to use asynchronous removal.</param>
	/// <returns>The asynchronous test.</returns>
	[TestCase(false)]
	[TestCase(true)]
	public async Task RemoveDuplicatesPreservesUniqueItems(bool useAsync)
	{
		using DuplicateTestData data = new();
		(string first, string copy, string unique) = CreateMixedSet(data);

		int reported = await data.RemoveDuplicatesAsync(useAsync, false)
			.ConfigureAwait(true);

		HashSet<string> survivors = data.GetItemIds();
		Assert.That(survivors, Has.Count.EqualTo(2));
		Assert.That(survivors, Does.Contain(unique));
		survivors.IntersectWith([first, copy]);
		Assert.That(survivors, Has.Count.EqualTo(1));
		Assert.That(data.GetDeletedCount(), Is.EqualTo(1));
		Assert.That(reported, Is.EqualTo(1));
	}

	/// <summary>Distinct mail-field boundaries must survive removal.</summary>
	/// <param name="useAsync">Whether to use asynchronous removal.</param>
	/// <returns>The asynchronous test.</returns>
	[TestCase(false)]
	[TestCase(true)]
	public async Task StringFieldBoundariesAreNotDeleted(bool useAsync)
	{
		using DuplicateTestData data = new();
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem second = DuplicateTestData.Copy(first);

		try
		{
			DuplicateTestData.SetBoundaryFields(first, "AB", "C");
			DuplicateTestData.SetBoundaryFields(second, "A", "BC");
			Assert.That(first.Categories, Is.EqualTo("AB"));
			Assert.That(first.Companies, Is.EqualTo("C"));
			string firstSynopsis = OutlookMail.GetSynopses(first);
			string secondSynopsis = OutlookMail.GetSynopses(second);
			Assert.That(secondSynopsis, Is.EqualTo(firstSynopsis));
			DuplicateTestData.RequiredHash(first);
			DuplicateTestData.RequiredHash(second);
		}
		finally
		{
			DuplicateTestData.Release(first, second);
		}

		HashSet<string> before = data.GetItemIds();
		Assert.That(before, Has.Count.EqualTo(2));

		await data.RemoveDuplicatesAsync(useAsync, false)
			.ConfigureAwait(true);

		HashSet<string> after = data.GetItemIds();
		Assert.That(
			after,
			Is.EquivalentTo(before),
			"Distinct field boundaries must not cause a false deletion.");
		Assert.That(data.GetDeletedCount(), Is.Zero);
	}

	private static (string First, string Copy, string Unique) CreateMixedSet(
		DuplicateTestData data)
	{
		Outlook.MailItem first = data.CreateMail();
		Outlook.MailItem copy = DuplicateTestData.Copy(first);
		Outlook.MailItem unique = DuplicateTestData.Copy(first);

		try
		{
			unique.Body = "Different content with the same synopsis.";
			unique.Save();
			string firstHash = DuplicateTestData.RequiredHash(first);
			string copyHash = DuplicateTestData.RequiredHash(copy);
			string uniqueHash = DuplicateTestData.RequiredHash(unique);
			Assert.That(copyHash, Is.EqualTo(firstHash));
			Assert.That(uniqueHash, Is.Not.EqualTo(firstHash));
			string firstSynopsis = OutlookMail.GetSynopses(first);
			string uniqueSynopsis = OutlookMail.GetSynopses(unique);
			Assert.That(uniqueSynopsis, Is.EqualTo(firstSynopsis));

			return (first.EntryID, copy.EntryID, unique.EntryID);
		}
		finally
		{
			DuplicateTestData.Release(first, copy, unique);
		}
	}

	private static void DeleteCandidate(
		DuplicateTestData data,
		string candidateId,
		string keeperSynopsis,
		bool useStore,
		bool dryRun)
	{
		if (useStore)
		{
			OutlookStore store = new(data.Session);
			OutlookItem.DeleteDuplicate(
				store, candidateId, keeperSynopsis, dryRun);
		}
		else
		{
			OutlookItem.DeleteDuplicate(
				data.NameSpace, candidateId, keeperSynopsis, dryRun);
		}
	}
}
