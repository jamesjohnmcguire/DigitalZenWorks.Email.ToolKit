/////////////////////////////////////////////////////////////////////////////
// <copyright file="EmlMessages.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Eml messages class.
	/// </summary>
	public static class EmlMessages
	{
		/// <summary>
		/// Get all eml or text files.
		/// </summary>
		/// <param name="location">The path location to check.</param>
		/// <returns>a list of eml and text files.</returns>
		public static IEnumerable<string> GetFiles(string location)
		{
			List<string> extensions = new () { ".eml", ".txt" };
			IEnumerable<string> allFiles =
				Directory.EnumerateFiles(location, "*.*");

			IEnumerable<string> query =
				allFiles.Where(file =>
					file.EndsWith(
						extensions[0], StringComparison.OrdinalIgnoreCase) ||
					file.EndsWith(
						extensions[1], StringComparison.OrdinalIgnoreCase));

			return query;
		}
	}
}
