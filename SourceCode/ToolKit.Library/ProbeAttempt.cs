/////////////////////////////////////////////////////////////////////////////
// <copyright file="ProbeAttempt.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Threading;
using global::Common.Logging;

/// <summary>
/// Keeps one worker's completion and failure isolated from later attempts.
/// </summary>
internal sealed class ProbeAttempt
{
	// Keep the coordinator's logger category for all probe diagnostics.
	private static readonly ILog Log =
		LogManager.GetLogger<OutlookApplicationProbe>();

	private readonly Action probeApplication;

	/// <summary>
	/// Initializes a new instance of the <see cref="ProbeAttempt"/> class.
	/// </summary>
	/// <param name="probeApplication">The STA operation to perform.</param>
	public ProbeAttempt(Action probeApplication)
	{
		this.probeApplication = probeApplication;
		Worker = new Thread(Run);
		Worker.SetApartmentState(ApartmentState.STA);
		Worker.IsBackground = true;
	}

	/// <summary>
	/// Gets the worker whose exit includes completion of cleanup.
	/// </summary>
	public Thread Worker { get; }

	/// <summary>
	/// Gets the failure, read by callers only after joining the worker.
	/// </summary>
	public ExceptionDispatchInfo? Failure { get; private set; }

	[SuppressMessage(
		"Design",
		"CA1031:Do not catch general exception types",
		Justification = "Transport worker failures; log unobserved late errors.")]
	private void Run()
	{
		try
		{
			Log.Debug("Outlook activation probe started.");
			probeApplication();
			Log.Debug("Outlook activation probe and cleanup completed.");
		}
		catch (Exception exception)
		{
			Failure = ExceptionDispatchInfo.Capture(exception);
			Log.Error(
				"Outlook activation probe failed during activation or " +
				"cleanup.", exception);
		}
	}
}
