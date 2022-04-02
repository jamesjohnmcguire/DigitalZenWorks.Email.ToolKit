/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookAccount.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Microsoft.Office.Interop.Outlook;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Represents an Outlook account.
	/// </summary>
	public class OutlookAccount
	{
		private readonly Application application;
		private readonly NameSpace session;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookAccount"/> class.
		/// </summary>
		public OutlookAccount()
		{
			application = new ();

			session = application.Session;
		}

		/// <summary>
		/// Gets the default session (Outlook namespace).
		/// </summary>
		/// <value>The default session (Outlook namespace).</value>
		public NameSpace Session { get { return session; } }
	}
}
