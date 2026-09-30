/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookFactoryTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using NUnit.Framework;

internal sealed class OutlookFactoryTests
{
	/// <summary>
	/// Verifies that IsOutlookAvailable returns true when the factory's
	/// CreateApplication operation succeeds.
	/// </summary>
	/// <remarks>Uses a TestOutlookFactory with CreateApplicationSucceeds set to
	/// true, calls IsOutlookAvailable with a timeout of 10, and asserts the
	/// result is true.</remarks>
	[Test]
	public void IsOutlookAvailableReturnsTrueWhenCreateApplicationSucceeds()
	{
		TestOutlookFactory factory = new();
		factory.CreateApplicationSucceeds = true;

		bool available = factory.CanCreateApplication(10);

		Assert.IsTrue(available);
	}

	/// <summary>
	/// Verifies that IsOutlookAvailable returns false when the factory's
	/// startup delay exceeds the provided timeout.
	/// </summary>
	/// <remarks>Initializes a TestOutlookFactory with DelayMilliseconds set to
	/// 2000 and calls IsOutlookAvailable(0), asserting a false result.
	/// </remarks>
	[Test]
	public void IsOutlookAvailableReturnsFalseOnTimeout()
	{
		TestOutlookFactory factory = new();
		factory.DelayMilliseconds = 2000;

		bool available = factory.CanCreateApplication(0);

		Assert.IsFalse(available);
	}
}
