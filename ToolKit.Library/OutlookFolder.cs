/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFolder.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Runtime.InteropServices;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Represents an Outlook Folder.
	/// </summary>
	internal class OutlookFolder
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// Add folder in safe context.
		/// </summary>
		/// <remarks>If there is a folder already existing with the given
		/// folder name, this method will return that folder.</remarks>
		/// <param name="parentFolder">The parent folder.</param>
		/// <param name="folderName">The new folder name.</param>
		/// <returns>The added or existing folder.</returns>
		public static MAPIFolder AddFolder(
			MAPIFolder parentFolder, string folderName)
		{
			MAPIFolder pstFolder = null;

			if (parentFolder != null && !string.IsNullOrWhiteSpace(folderName))
			{
				pstFolder = GetSubFolder(parentFolder, folderName);

				if (pstFolder == null)
				{
					Log.Info("Adding outlook folder: " + folderName);

					try
					{
						pstFolder = parentFolder.Folders.Add(folderName);
					}
					catch (COMException exception)
					{
						Log.Warn(exception.ToString());
					}
				}
			}

			return pstFolder;
		}

		/// <summary>
		/// Get the folder's full path.
		/// </summary>
		/// <param name="folder">The folder to check.</param>
		/// <returns>The folder's full path.</returns>
		public static string GetFolderPath(MAPIFolder folder)
		{
			string path = null;

			if (folder != null)
			{
				path = folder.Name;
				MAPIFolder parent = folder.Parent;

				while (parent is not null && parent is MAPIFolder)
				{
					path = parent.Name + "/" + path;
					folder = parent;

					if (folder.Parent is not null &&
						folder.Parent is MAPIFolder)
					{
						parent = folder.Parent;
					}
					else
					{
						parent = null;
					}
				}

				string storeName = OutlookStorage.GetStoreName(folder.Store);
				path = storeName + "::" + path;
			}

			return path;
		}

		/// <summary>
		/// Get sub folder from parent.
		/// </summary>
		/// <param name="parentFolder">The parent folder.</param>
		/// <param name="folderName">The new folder name.</param>
		/// <returns>The added folder.</returns>
		public static MAPIFolder GetSubFolder(
			MAPIFolder parentFolder, string folderName)
		{
			MAPIFolder pstFolder = null;

			if (parentFolder != null && !string.IsNullOrWhiteSpace(folderName))
			{
				int total = parentFolder.Folders.Count;

				for (int index = 1; index <= total; index++)
				{
					MAPIFolder subFolder = parentFolder.Folders[index];

					if (folderName.Equals(
						subFolder.Name, StringComparison.OrdinalIgnoreCase))
					{
						pstFolder = subFolder;
						break;
					}

					Marshal.ReleaseComObject(subFolder);
				}
			}

			return pstFolder;
		}
	}
}
