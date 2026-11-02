/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookApplicationProbe.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit;

using System;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using global::Common.Logging;

/// <summary>
/// Coordinates one outstanding STA probe without transferring COM objects.
/// The operation includes both activation and cleanup on the worker thread.
/// </summary>
internal sealed class OutlookApplicationProbe
{
	private static readonly ILog Log =
		LogManager.GetLogger<OutlookApplicationProbe>();

	private readonly object syncRoot = new();
	private readonly Action probeApplication;
	private ProbeAttempt? activeAttempt;

	/// <summary>
	/// Initializes a new instance of the
	/// <see cref="OutlookApplicationProbe"/> class.
	/// </summary>
	/// <param name="probeApplication">Activation and cleanup to execute
	/// entirely on the temporary STA thread.</param>
	public OutlookApplicationProbe(Action probeApplication)
	{
		this.probeApplication = probeApplication ??
			throw new ArgumentNullException(nameof(probeApplication));
	}

	/// <summary>
	/// Starts a probe or joins the outstanding attempt for the requested wait.
	/// </summary>
	/// <param name="timeOutSeconds">Probe wait in seconds, from 0 to 2147483.
	/// Zero starts or joins an attempt without waiting for completion.</param>
	/// <returns>Whether the joined attempt completed successfully.</returns>
	public bool CanCreateApplication(int timeOutSeconds)
	{
		// Validate before starting work. Thread.Join accepts at most
		// Int32.MaxValue milliseconds; there is no infinite-wait option here.
		if (timeOutSeconds < 0 || timeOutSeconds > int.MaxValue / 1000)
		{
			throw new ArgumentOutOfRangeException(
				nameof(timeOutSeconds),
				timeOutSeconds,
				"The probe wait must be between 0 and 2147483 seconds.");
		}

		TimeSpan timeout = TimeSpan.FromSeconds(timeOutSeconds);
		ProbeAttempt attempt;

		lock (syncRoot)
		{
			// A timed-out worker remains active until it actually exits.
			// Its cleanup must finish before another probe can start.
			if (activeAttempt == null || activeAttempt.Worker.IsAlive == false)
			{
				ProbeAttempt nextAttempt = new(probeApplication);
				nextAttempt.Worker.Start();
				activeAttempt = nextAttempt;
			}

			attempt = activeAttempt;
		}

		// Do not hold the coordinator lock while waiting or executing COM.
		// Each caller keeps its own attempt reference, so a later attempt
		// cannot overwrite the result it is about to inspect.
		bool finished = attempt.Worker.Join(timeout);
		bool isAvailable = false;

		if (finished == false)
		{
			Log.Warn("Outlook activation probe wait timed out. " +
				"The worker is still running; retries share that attempt.");
		}
		else
		{
			ExceptionDispatchInfo? failure = attempt.Failure;

			if (failure == null)
			{
				isAvailable = true;
			}
			else if (failure.SourceException is not COMException &&
				failure.SourceException is not InvalidOperationException)
			{
				// Preserve unexpected failures for callers that are still
				// waiting. Late failures have also been logged by the worker.
				failure.Throw();
			}
		}

		return isAvailable;
	}
}
