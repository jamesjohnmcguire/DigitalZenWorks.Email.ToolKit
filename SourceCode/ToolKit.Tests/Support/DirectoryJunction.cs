/////////////////////////////////////////////////////////////////////////////
// <copyright file="DirectoryJunction.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Diagnostics;
using System.IO;

/// <summary>
/// Provides a temporary directory alias without symbolic-link privileges.
/// </summary>
internal sealed class DirectoryJunction : IDisposable
{
	private readonly DirectoryInfo directory;

	/// <summary>
	/// Initializes a new instance of the <see cref="DirectoryJunction"/> class.
	/// </summary>
	/// <param name="target">The existing directory to alias.</param>
	public DirectoryJunction(string target)
	{
		directory = Directory.CreateTempSubdirectory("Email.ToolKit.Link.");
		Path = System.IO.Path.Combine(directory.FullName, "alias");

		ProcessStartInfo startInfo = new("cmd.exe");
		startInfo.Arguments = $"/c mklink /J \"{Path}\" \"{target}\"";
		startInfo.CreateNoWindow = true;
		startInfo.UseShellExecute = false;
		startInfo.RedirectStandardOutput = true;
		startInfo.RedirectStandardError = true;

		using Process process = Process.Start(startInfo)!;
		string output = process.StandardOutput.ReadToEnd();
		string error = process.StandardError.ReadToEnd();
		process.WaitForExit();

		if (process.ExitCode != 0)
		{
			directory.Delete();
			throw new IOException("Cannot create test junction: " +
				output + error);
		}
	}

	/// <summary>
	/// Gets the path through which the target directory can be accessed.
	/// </summary>
	public string Path { get; }

	/// <summary>
	/// Removes the alias without deleting its target or the target's files.
	/// </summary>
	public void Dispose()
	{
		Directory.Delete(Path);
		directory.Delete();
	}
}
