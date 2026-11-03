/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandExitCodeManualTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;
using System.Diagnostics;
using System.Threading;
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
	/// <returns>A task that completes after the exit code is checked.</returns>
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
		using CancellationTokenSource timeout = new(TimeSpan.FromSeconds(30));

		try
		{
			await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
		}
		catch (OperationCanceledException) when (timeout.IsCancellationRequested)
		{
			if (process.HasExited == false)
			{
				process.Kill();
			}

			await process.WaitForExitAsync().ConfigureAwait(false);
			Assert.Fail("Det did not exit within 30 seconds.");
		}

		string output = await outputTask.ConfigureAwait(false);
		string error = await errorTask.ConfigureAwait(false);
		string combinedOutput = output + error;

		Assert.That(combinedOutput, Does.Contain("Unable to Continue"));
		Assert.That(process.ExitCode, Is.Not.Zero);
	}
}
