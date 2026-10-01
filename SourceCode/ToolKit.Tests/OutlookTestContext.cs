/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookTestContext.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Owns test data and COM references acquired by an Outlook integration test.
/// </summary>
/// <remarks>
/// Releases its own references without quitting a potentially shared Outlook.
/// PST files may remain locked by Outlook after they have been detached.
/// </remarks>
internal sealed class OutlookTestContext : IDisposable
{
	private readonly List<object> references = new();
	private readonly HashSet<string> storePaths =
		new(StringComparer.OrdinalIgnoreCase);

	private bool disposed;

	/// <summary>
	/// Initializes a new instance of the <see cref="OutlookTestContext"/> class.
	/// </summary>
	public OutlookTestContext()
	{
		DirectoryPath =
			Directory.CreateTempSubdirectory("Email.ToolKit.Outlook.").FullName;
		Application = Track(new Outlook.Application());
		NameSpace = Track(Application.Session);
	}

	/// <summary>
	/// Gets the real Outlook application.
	/// </summary>
	public Outlook.Application Application { get; }

	/// <summary>
	/// Gets the namespace used to independently inspect test outcomes.
	/// </summary>
	public Outlook.NameSpace NameSpace { get; }

	/// <summary>
	/// Gets the directory containing only this test's temporary data.
	/// </summary>
	public string DirectoryPath { get; }

	/// <summary>
	/// Tracks one acquired COM reference for release during cleanup.
	/// </summary>
	/// <typeparam name="T">The acquired COM interface.</typeparam>
	/// <param name="reference">The reference to retain.</param>
	/// <returns>The supplied reference.</returns>
	public T Track<T>(T reference)
		where T : class
	{
		references.Add(reference);
		return reference;
	}

	/// <summary>
	/// Registers a test-owned PST path before attempting to attach it.
	/// </summary>
	/// <param name="fileName">The PST file name.</param>
	/// <returns>The absolute PST path.</returns>
	public string StorePath(string fileName)
	{
		string path = Path.Combine(DirectoryPath, fileName);
		storePaths.Add(path);
		return path;
	}

	/// <summary>
	/// Counts matching stores directly through Outlook, without the library.
	/// </summary>
	/// <param name="path">The absolute PST path.</param>
	/// <returns>The number of matching stores.</returns>
	public int CountStores(string path)
	{
		int count = 0;
		Outlook.Stores stores = NameSpace.Stores;

		try
		{
			for (int index = 1; index <= stores.Count; index++)
			{
				Outlook.Store store = stores[index];

				try
				{
					if (string.Equals(
						store.FilePath, path, StringComparison.OrdinalIgnoreCase))
					{
						count++;
					}
				}
				finally
				{
					Marshal.ReleaseComObject(store);
				}
			}
		}
		finally
		{
			Marshal.ReleaseComObject(stores);
		}

		return count;
	}

	/// <summary>
	/// Detaches only this test's PSTs and releases its acquired references.
	/// </summary>
	public void Dispose()
	{
		if (disposed)
		{
			return;
		}

		disposed = true;

		try
		{
			Outlook.Stores stores = NameSpace.Stores;

			try
			{
				for (int index = stores.Count; index >= 1; index--)
				{
					Outlook.Store store = stores[index];

					try
					{
						if (storePaths.Contains(store.FilePath))
						{
							Outlook.MAPIFolder root = store.GetRootFolder();

							try
							{
								NameSpace.RemoveStore(root);
							}
							finally
							{
								Marshal.ReleaseComObject(root);
							}
						}
					}
					finally
					{
						Marshal.ReleaseComObject(store);
					}
				}
			}
			finally
			{
				Marshal.ReleaseComObject(stores);
			}
		}
		finally
		{
			for (int index = references.Count - 1; index >= 0; index--)
			{
				Marshal.ReleaseComObject(references[index]);
			}

			references.Clear();

			try
			{
				Directory.Delete(DirectoryPath, true);
			}
			catch (IOException exception)
			{
				TestContext.Progress.WriteLine(
					"Outlook retained a lock on test data at " +
					DirectoryPath + ": " + exception.Message);
			}
		}
	}
}
