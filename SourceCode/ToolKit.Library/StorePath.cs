/////////////////////////////////////////////////////////////////////////////
// <copyright file="StorePath.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

/// <summary>
/// Compares store paths, including aliases in ancestor directories.
/// </summary>
internal static class StorePath
{
	/// <summary>
	/// Determines whether paths name the same store through directory links.
	/// </summary>
	/// <param name="firstPath">The first store file path.</param>
	/// <param name="secondPath">The second store file path.</param>
	/// <returns>True for matching paths or resolved Windows paths.</returns>
	internal static bool AreEqual(string firstPath, string secondPath)
	{
		if (firstPath.Equals(secondPath, StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}

		if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
		{
			return false;
		}

		string? firstResolved = GetResolvedPath(firstPath);
		if (firstResolved == null)
		{
			return false;
		}

		string? secondResolved = GetResolvedPath(secondPath);

		return firstResolved.Equals(
			secondResolved, StringComparison.OrdinalIgnoreCase);
	}

	private static string? GetResolvedPath(string path)
	{
		// Query metadata only, without reading or locking Outlook's PST data.
		using SafeFileHandle handle = NativeMethods.CreateFileW(
			path,
			0,
			FileShare.ReadWrite | FileShare.Delete,
			0,
			FileMode.Open,
			0,
			0);

		if (handle.IsInvalid)
		{
			return null;
		}

		char[] buffer = new char[260];
		uint length = NativeMethods.GetFinalPathNameByHandleW(
			handle, buffer, (uint)buffer.Length, 0);

		if (length >= buffer.Length)
		{
			buffer = new char[length];
			length = NativeMethods.GetFinalPathNameByHandleW(
				handle, buffer, (uint)buffer.Length, 0);
		}

		if (length == 0 || length >= buffer.Length)
		{
			// Unavailable metadata must not imply that two paths are equal.
			return null;
		}

		return new string(buffer, 0, (int)length);
	}
}
