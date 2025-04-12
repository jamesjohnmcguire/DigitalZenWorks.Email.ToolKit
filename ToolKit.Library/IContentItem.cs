/////////////////////////////////////////////////////////////////////////////
// <copyright file="IContentItem.cs" company="James John McGuire">
// Copyright © 2021 - 2025 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Microsoft.Office.Interop.Outlook;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Content Item Interface.
	/// </summary>
	public interface IContentItem
	{
		/// <summary>
		/// Gets the item's hash text.
		/// </summary>
		/// <value>The item's hash stext.</value>
		string Hash { get; }

		/// <summary>
		/// Gets the item's synopses text.
		/// </summary>
		/// <value>The item's synopses text.</value>
		string Synopses { get; }

		/// <summary>
		/// Deletes the given item.
		/// </summary>
		public void Delete();

		/// <summary>
		/// Moves the given item.
		/// </summary>
		/// <param name="destination">The destination folder.</param>
		public void Move(MAPIFolder destination);
	}
}
