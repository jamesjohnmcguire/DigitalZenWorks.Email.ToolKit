/////////////////////////////////////////////////////////////////////////////
// <copyright file="SuccessFactory.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

namespace DigitalZenWorks.Email.ToolKit.Tests;

public sealed class SuccessFactory : IOutlookFactory
{
	public bool IsOutlookAvailable(int timeOutSeconds)
	{
		return true;
	}
}
