/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookAccount.cs" company="James John McGuire">
// Copyright © 2021 - 2025 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Threading.Tasks;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Represents an Outlook account.
	/// </summary>
	public sealed class OutlookAccount
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private static readonly OutlookAccount InternalInstance = new ();

		private readonly Application application;
		private readonly NameSpace session;

		// Explicit static constructor to tell C# compiler
		// not to mark type as beforefieldinit
		static OutlookAccount()
		{
		}

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookAccount"/> class.
		/// </summary>
		private OutlookAccount()
		{
			application = new ();

			session = application.Session;
		}

		/// <summary>
		/// Gets the singleton instance of this class.
		/// </summary>
		/// <value>The singleton instance of this class.</value>
		public static OutlookAccount Instance
		{
			get { return InternalInstance; }
		}

		/// <summary>
		/// Gets the Outlook application object.
		/// </summary>
		/// <value>The Outlook application object.</value>
		public Application Application { get { return application; } }

		/// <summary>
		/// Gets the default session (Outlook namespace).
		/// </summary>
		/// <value>The default session (Outlook namespace).</value>
		public NameSpace Session { get { return session; } }

		/// <summary>
		/// Create mail item.
		/// </summary>
		/// <param name="recipient">The recipient of the mail.</param>
		/// <param name="subject">The subject of the mail.</param>
		/// <param name="body">The body of the mail.</param>
		/// <returns>The created mail item.</returns>
		public MailItem CreateMailItem(
			string recipient, string subject, string body)
		{
			MailItem mailItem =
				(MailItem)application.CreateItem(OlItemType.olMailItem);

			mailItem.Display(false);

			mailItem.To = recipient;
			mailItem.Subject = subject;
			mailItem.Body = body;

			return mailItem;
		}

		/// <summary>
		/// Empty deleted items folder.
		/// </summary>
		public void EmptyDeletedItemsFolder()
		{
			Store store = session.DefaultStore;

			OutlookStore.EmptyDeletedItemsFolder(store);
		}

		/// <summary>
		/// Create a new pst storage file.
		/// </summary>
		/// <param name="path">The path to the pst file.</param>
		/// <returns>A store object.</returns>
		public Store GetStore(string path)
		{
			Store newPst = null;

			path = Path.GetFullPath(path);

			string extension = Path.GetExtension(path);

			if (!extension.Equals(".pst", StringComparison.OrdinalIgnoreCase))
			{
				// Attempt to fix mistaken or missing file extension.
				path += ".pst";
			}

			// If the .pst file does not exist, Microsoft Outlook creates it.
			session.AddStore(path);

			int total = session.Stores.Count;

			for (int index = 1; index <= total; index++)
			{
				Store store = null;

				try
				{
					store = session.Stores[index];
				}
				catch (UnauthorizedAccessException exception)
				{
					Log.Error(exception.ToString());
				}

				if (store == null)
				{
					Log.Warn("Enumerating stores - store is null");
				}
				else
				{
					string filePath = store.FilePath;

					if (!string.IsNullOrWhiteSpace(filePath) &&
						filePath.Equals(
							path, StringComparison.OrdinalIgnoreCase))
					{
						newPst = store;
						break;
					}
				}
			}

			if (newPst == null)
			{
				Log.Warn("Store not found: " + path);
			}

			return newPst;
		}

		/// <summary>
		/// Removes a store from Outlook.
		/// </summary>
		/// <param name="path">The path to the pst file.</param>
		/// <returns>remove result.</returns>
		public bool RemoveStore(string path)
		{
			bool result = false;

			Log.Info("Begin to Removing store: " + path);

			path = Path.GetFullPath(path);
			string extension = Path.GetExtension(path);

			if (!extension.Equals(".pst", StringComparison.OrdinalIgnoreCase))
			{
				// Attempt to fix mistaken or missing file extension.
				path += ".pst";
			}

			Store store = GetStore(path);

			if (store != null)
			{
				MAPIFolder rootFolder = store.GetRootFolder();
				session.RemoveStore(rootFolder);

				Log.Info("Store removed successfully: " + path);
				result = true;
			}
			else
			{
				Log.Warn("Store not found: " + path);
			}

			return result;
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		public void MergeFolders(bool dryRun)
		{
			OutlookStore outlookStorage = new (this);
			uint totalFolders = 0;
			int totalStores = session.Stores.Count;

			for (int index = 1; index <= totalStores; index++)
			{
				Store store = session.Stores[index];

				totalFolders += outlookStorage.MergeFolders(store, dryRun);
			}

			Log.Info("Remove empty folder complete - total folders checked: " +
				totalFolders);
		}

		/// <summary>
		/// Merge duplicate folders.
		/// </summary>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task MergeFoldersAsync(bool dryRun)
		{
			OutlookStore outlookStorage = new (this);
			uint totalFolders = 0;
			int totalStores = session.Stores.Count;

			for (int index = 1; index <= totalStores; index++)
			{
				Store store = session.Stores[index];

				totalFolders += await outlookStorage.MergeFoldersAsync(
					store, dryRun).ConfigureAwait(false);
			}

			Log.Info("Remove empty folder complete - total folders checked: " +
				totalFolders);
		}

		/// <summary>
		/// Remove duplicates items from default account.
		/// </summary>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <param name="flush">Indicates whether to empty the deleted items
		/// folder.</param>
		public void RemoveDuplicates(bool dryRun, bool flush)
		{
			OutlookStore outlookStorage = new (this);
			int total = session.Stores.Count;

			for (int index = 1; index <= total; index++)
			{
				Store store = session.Stores[index];

				outlookStorage.RemoveDuplicates(store, dryRun, flush);
			}
		}

		/// <summary>
		/// Remove duplicates items from default account.
		/// </summary>
		/// <param name="dryRun">Indicates whether this is a 'dry run'
		/// or not.</param>
		/// <param name="flush">Indicates whether to empty the deleted items
		/// folder.</param>
		/// <returns>A <see cref="Task"/> representing the asynchronous
		/// operation.</returns>
		public async Task RemoveDuplicatesAsync(bool dryRun, bool flush)
		{
			OutlookStore outlookStorage = new (this);
			int total = session.Stores.Count;

			for (int index = 1; index <= total; index++)
			{
				Store store = session.Stores[index];

				await outlookStorage.RemoveDuplicatesAsync(
					store, dryRun, flush).ConfigureAwait(false);
			}
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <returns>The count of removed folders.</returns>
		public int RemoveEmptyFolders()
		{
			int total = session.Stores.Count;
			int removedFolders = 0;

			for (int index = 1; index <= total; index++)
			{
				Store store = session.Stores[index];

				removedFolders += OutlookStore.RemoveEmptyFolders(store);
			}

			return removedFolders;
		}

		/// <summary>
		/// Remove all empty folders.
		/// </summary>
		/// <returns>The count of removed folders.</returns>
		public async Task<int> RemoveEmptyFoldersAsync()
		{
			int total = session.Stores.Count;
			int removedFolders = 0;

			for (int index = 1; index <= total; index++)
			{
				Store store = session.Stores[index];

				removedFolders += await OutlookStore.RemoveEmptyFoldersAsync(
					store).ConfigureAwait(false);
			}

			return removedFolders;
		}
	}
}
