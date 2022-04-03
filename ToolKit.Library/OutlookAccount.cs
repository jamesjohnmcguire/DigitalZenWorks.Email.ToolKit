/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookAccount.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using Microsoft.Office.Interop.Outlook;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Represents an Outlook account.
	/// </summary>
	public sealed class OutlookAccount
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private static readonly OutlookAccount InternalInstance = new ();

		private readonly Application application;
		private readonly NameSpace session;

		// Explicit static constructor to tell C# compiler
		// not to mark type as beforefieldinit
		static OutlookAccount()
		{
		}

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="OutlookAccount"/> class.
		/// </summary>
		private OutlookAccount()
		{
			application = new ();

			session = application.Session;
		}

		/// <summary>
		/// Gets the singleton instance of this class.
		/// </summary>
		/// <value>The singleton instance of this class.</value>
		public static OutlookAccount Instance
		{
			get { return InternalInstance; }
		}

		/// <summary>
		/// Gets the Outlook application object.
		/// </summary>
		/// <value>The Outlook application object.</value>
		public Application Application { get { return application; } }

		/// <summary>
		/// Gets the default session (Outlook namespace).
		/// </summary>
		/// <value>The default session (Outlook namespace).</value>
		public NameSpace Session { get { return session; } }

		/// <summary>
		/// Create mail item.
		/// </summary>
		/// <param name="recipient">The recipient of the mail.</param>
		/// <param name="subject">The subject of the mail.</param>
		/// <param name="body">The body of the mail.</param>
		/// <returns>The created mail item.</returns>
		public MailItem CreateMailItem(
			string recipient, string subject, string body)
		{
			MailItem mailItem =
				(MailItem)application.CreateItem(OlItemType.olMailItem);

			mailItem.Display(false);

			mailItem.To = recipient;
			mailItem.Subject = subject;
			mailItem.Body = body;

			return mailItem;
		}
	}
}
