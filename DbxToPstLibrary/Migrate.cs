/////////////////////////////////////////////////////////////////////////////
// <copyright file="Migrate.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using DigitalZenWorks.Email.DbxOutlookExpress;
using System;
using System.Globalization;
using System.IO;

[assembly: CLSCompliant(true)]

namespace DbxToPstLibrary
{
	/// <summary>
	/// Migrate Dbx to Pst class.
	/// </summary>
	public static class Migrate
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// Dbx directory to pst.
		/// </summary>
		/// <param name="directoryPath">The directory path to migrate.</param>
		public static void DbxDirectoryToPst(string directoryPath)
		{
			DbxSet dbxSet = new (directoryPath);
			dbxSet.Migrate();
		}

		/// <summary>
		/// Dbx files to pst.
		/// </summary>
		/// <param name="filePath">The file path to migrate.</param>
		public static void DbxFileToPst(string filePath)
		{
			DbxSet dbxSet = new (filePath);
		}

		/// <summary>
		/// Dbx to pst.
		/// </summary>
		/// <param name="path">the path of the dbx element.</param>
		/// <returns>A value indicating success or not.</returns>
		public static bool DbxToPst(string path)
		{
			bool result = false;

			if (Directory.Exists(path))
			{
				DbxDirectoryToPst(path);
				result = true;
			}
			else if (File.Exists(path))
			{
				DbxFileToPst(path);
				result = true;
			}
			else
			{
				Log.Error("Invalid path");
			}

			return result;
		}
	}
}
