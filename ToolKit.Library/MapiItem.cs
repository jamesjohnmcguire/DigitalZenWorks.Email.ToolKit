/////////////////////////////////////////////////////////////////////////////
// <copyright file="MapiItem.cs" company="James John McGuire">
// Copyright © 2021 - 2025 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Microsoft.Office.Interop.Outlook;
using System;
using System.Threading.Tasks;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Provides comparision support for Outlook MAPI items.
	/// </summary>
	public static class MapiItem
	{
		/// <summary>
		/// Delete duplicate item.
		/// </summary>
		/// <param name="session">The Outlook session.</param>
		/// <param name="duplicateId">The duplicate id.</param>
		/// <param name="keeperSynopses">The keeper synopses.</param>
		/// <param name="dryRun">Indicates if this is a dry run or not.</param>
		[Obsolete("DeleteDuplicate is deprecated, " +
			"please use OutlookItem.DeleteDuplicate instead.")]
		public static void DeleteDuplicate(
			NameSpace session,
			string duplicateId,
			string keeperSynopses,
			bool dryRun)
		{
			OutlookItem.DeleteDuplicate(
				session, duplicateId, keeperSynopses, dryRun);
		}

		/// <summary>
		/// Deletes the given item.
		/// </summary>
		/// <param name="item">The item to delete.</param>
		[Obsolete("DeleteItem is deprecated, " +
			"please use OutlookItem.Delete instead.")]
		public static void DeleteItem(object item)
		{
			OutlookItem.Delete(item);
		}

		/// <summary>
		/// Gets the item's hash.
		/// </summary>
		/// <param name="mapiItem">The items to compute.</param>
		/// <returns>The item's hash encoded in base 64.</returns>
		[Obsolete("GetItemHash is deprecated, " +
			"please use OutlookItem.Hash instead.")]
		public static string GetItemHash(object mapiItem)
		{
			OutlookItem outlookItem = new (mapiItem);

			string hashBase64 = outlookItem.Hash;

			return hashBase64;
		}

		/// <summary>
		/// Gets the item's hash.
		/// </summary>
		/// <param name="mapiItem">The items to compute.</param>
		/// <returns>The item's hash encoded in base 64.</returns>
		[Obsolete("GetItemHashAsync is deprecated, " +
			"please use OutlookItem.GetHashAsync instead.")]
		public static async Task<string> GetItemHashAsync(object mapiItem)
		{
			string hashBase64 = await OutlookItem.GetHashAsync(mapiItem).
										ConfigureAwait(false);

			return hashBase64;
		}

		/// <summary>
		/// Get the item's synopses.
		/// </summary>
		/// <param name="mapiItem">The specific MAPI item to check.</param>
		/// <returns>The synoses of the item.</returns>
		[Obsolete("GetItemSynopses is deprecated, " +
			"please use OutlookItem.Synopses instead.")]
		public static string GetItemSynopses(object mapiItem)
		{
			OutlookItem outlookItem = new (mapiItem);
			string synopses = outlookItem.Synopses;

			return synopses;
		}

		/// <summary>
		/// Get the current item's folder path.
		/// </summary>
		/// <param name="mapiItem">The item to check.</param>
		/// <returns>The current item's folder path.</returns>
		[Obsolete("GetPath is deprecated, " +
			"please use OutlookItem.GetPath instead.")]
		public static string GetPath(object mapiItem)
		{
			string path = OutlookItem.GetPath(mapiItem);

			return path;
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="item">The item to move.</param>
		/// <param name="destination">The destination folder.</param>
		[Obsolete("MoveItem is deprecated, " +
			"please use OutlookItem.Move instead.")]
		public static void MoveItem(object item, MAPIFolder destination)
		{
			OutlookItem outlookItem = new (item);

			outlookItem.Move(destination);
		}

		/// <summary>
		/// Move item to destination folder.
		/// </summary>
		/// <param name="item">The item to move.</param>
		/// <param name="destination">The destination folder.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		[Obsolete("MoveItemAsync is deprecated, " +
			"please use OutlookItem.MoveAsync instead.")]
		public static async Task MoveItemAsync(
			object item, MAPIFolder destination)
		{
			await OutlookItem.MoveAsync(item, destination).
				ConfigureAwait(false);
		}

		/// <summary>
		/// Removes the MimeOLE version number.
		/// </summary>
		/// <param name="header">The header to check.</param>
		/// <returns>The modified header.</returns>
		[Obsolete("RemoveMimeOleVersion is deprecated, " +
			"please use OutlookItem.RemoveMimeOleVersion instead.")]
		public static string RemoveMimeOleVersion(string header)
		{
			header = OutlookItem.RemoveMimeOleVersion(header);

			return header;
		}
	}
}
