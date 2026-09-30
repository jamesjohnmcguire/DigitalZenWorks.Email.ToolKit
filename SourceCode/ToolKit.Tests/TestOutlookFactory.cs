/////////////////////////////////////////////////////////////////////////////
// <copyright file="TestOutlookFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

/// <summary>
/// Test implementation of OutlookFactory that simulates application creation
/// success or failure and supports a configurable delay.
/// </summary>
/// <remarks>Intended for use in unit tests to control whether application
/// creation succeeds and to simulate latency via DelayMilliseconds.</remarks>
internal sealed class TestOutlookFactory : OutlookFactory
{
	/// <summary>
	/// Gets or sets a value indicating whether true when creating an
	/// application succeeds; otherwise false.
	/// </summary>
	/// <remarks>Used in tests to simulate the outcome of an application
	/// creation operation.</remarks>
	public bool CreateApplicationSucceeds { get; set; }

	/// <summary>
	/// Gets or sets the delay in milliseconds.
	/// </summary>
	/// <remarks>Used in tests to simulate delays in application creation.
	/// </remarks>
	public int DelayMilliseconds { get; set; }
}
