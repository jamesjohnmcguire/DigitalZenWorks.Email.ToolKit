/////////////////////////////////////////////////////////////////////////////
// <copyright file="NativeMethods.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

/// <summary>
/// Declares Windows filesystem entry points used by store path lookup.
/// </summary>
internal static class NativeMethods
{
	/// <summary>
	/// Opens an existing file for metadata access.
	/// </summary>
	/// <param name="fileName">The existing file path.</param>
	/// <param name="desiredAccess">The access rights requested.</param>
	/// <param name="shareMode">The permitted sharing modes.</param>
	/// <param name="securityAttributes">Optional security
	/// attributes.</param>
	/// <param name="creationDisposition">The file opening mode.</param>
	/// <param name="flagsAndAttributes">Optional file flags.</param>
	/// <param name="templateFile">An optional template handle.</param>
	/// <returns>A handle closed automatically after the query.</returns>
	[DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
	[DllImport(
		"kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
	internal static extern SafeFileHandle CreateFileW(
		string fileName,
		uint desiredAccess,
		FileShare shareMode,
		nint securityAttributes,
		FileMode creationDisposition,
		uint flagsAndAttributes,
		nint templateFile);

	/// <summary>
	/// Gets the final path after resolving filesystem links.
	/// </summary>
	/// <param name="file">The open file handle.</param>
	/// <param name="path">The destination buffer.</param>
	/// <param name="capacity">The buffer size in characters.</param>
	/// <param name="flags">The path format flags.</param>
	/// <returns>The length, required capacity, or zero on
	/// failure.</returns>
	[DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
	[DllImport(
		"kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
	internal static extern uint GetFinalPathNameByHandleW(
		SafeFileHandle file, [Out] char[] path, uint capacity, uint flags);
}
