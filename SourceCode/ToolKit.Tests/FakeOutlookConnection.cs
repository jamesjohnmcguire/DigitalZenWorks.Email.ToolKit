/////////////////////////////////////////////////////////////////////////////
// <copyright file="FakeOutlookConnection.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System;

/// <summary>
/// A fake implementation of the IOutlookConnection interface for testing
/// purposes.
/// </summary>
internal sealed class FakeOutlookConnection
	: IOutlookConnection, IDisposable
{
	private readonly IOutlookSession? session;

	/// <summary>
	/// Initializes a new instance of the <see cref="FakeOutlookConnection"/>
	/// class.
	/// </summary>
	public FakeOutlookConnection()
		: this(new FakeOutlookSession())
	{
	}

	/// <summary>
	/// Initializes a new instance of the <see cref="FakeOutlookConnection"/>
	/// class.
	/// </summary>
	/// <param name="session">The Outlook session.</param>
	public FakeOutlookConnection(IOutlookSession? session)
	{
		this.session = session;
	}

	/// <summary>
	/// Gets the Outlook session.
	/// </summary>
	public IOutlookSession? Session
	{
		get
		{
			SessionAccessCount++;

			if (SessionException != null)
			{
				throw SessionException;
			}

			return session;
		}
	}

	/// <summary>
	/// Gets the number of session acquisitions.
	/// </summary>
	public int SessionAccessCount { get; private set; }

	/// <summary>
	/// Gets the number of quit requests.
	/// </summary>
	public int QuitCallCount { get; private set; }

	/// <summary>
	/// Gets or sets the failure to inject into session acquisition.
	/// </summary>
	public Exception? SessionException { get; set; }

	/// <summary>
	/// Gets or sets the failure to inject into disposal.
	/// </summary>
	public Exception? DisposeException { get; set; }

	/// <summary>
	/// Gets the number of disposal requests.
	/// </summary>
	public int DisposeCallCount { get; private set; }

	/// <summary>
	/// Records cleanup separately from explicit application shutdown.
	/// </summary>
	public void Dispose()
	{
		GC.SuppressFinalize(this);
		DisposeCallCount++;

		if (DisposeException != null)
		{
			throw DisposeException;
		}
	}

	/// <summary>
	/// Records an explicit application shutdown request.
	/// </summary>
	public void Quit()
	{
		QuitCallCount++;
	}
}
