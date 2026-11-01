/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandExitCodeManualTests.cs" company="James John McGuire">
// Copyright c 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Diagnostics;
using System.Threading.Tasks;
using NUnit.Framework;

/// <summary>
/// Checks the shipped command-line executable in an Outlook-unavailable host.
/// </summary>
[Category("ManualOutlook")]
internal sealed class CommandExitCodeManualTests
{
	/// <summary>
	/// Set EMAIL_TOOLKIT_DET_PATH to Det.exe and run on a Windows host where
	/// Outlook activation fails. No real PST paths are needed: connection must
	/// fail before the merge starts.
	/// </summary>
	[Test]
	[Explicit("Requires Det.exe on a host where Outlook cannot activate.")]
	public async Task MergeStoresReturnsFailureWhenOutlookCannotConnect()
	{
		string? executablePath = Environment.GetEnvironmentVariable(
			"EMAIL_TOOLKIT_DET_PATH");
		Assert.That(executablePath, Is.Not.Null.And.Not.Empty);

		ProcessStartInfo startInfo = new(executablePath!);
		startInfo.UseShellExecute = false;
		startInfo.RedirectStandardOutput = true;
		startInfo.RedirectStandardError = true;
		startInfo.ArgumentList.Add("merge-stores");
		startInfo.ArgumentList.Add("source.pst");
		startInfo.ArgumentList.Add("destination.pst");

		using Process process = Process.Start(startInfo)!;
		Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
		Task<string> errorTask = process.StandardError.ReadToEndAsync();
		bool exited = process.WaitForExit(30000);

		if (exited == false)
		{
			process.Kill();
			Assert.Fail("Det did not exit within 30 seconds.");
		}

		string output = await outputTask.ConfigureAwait(false);
		string error = await errorTask.ConfigureAwait(false);
		string combinedOutput = output + error;

		Assert.That(combinedOutput, Does.Contain("Unable to Continue"));
		Assert.That(process.ExitCode, Is.Not.Zero);
	}
}
