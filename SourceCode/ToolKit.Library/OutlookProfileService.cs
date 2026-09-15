/////////////////////////////////////////////////////////////////////////////
// <copyright file="OutlookProfileService.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace ToolKit.Library;

using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

/// <summary>
/// Represents a service for retrieving Outlook profile information.
/// </summary>
public class OutlookProfileService : IOutlookProfileService
{
	private readonly Application? application;

	/// <summary>
	/// Initializes a new instance of the
	/// <see cref="OutlookProfileService"/> class.
	/// </summary>
	/// <param name="application">The Outlook application instance.</param>
	public OutlookProfileService(Application? application)
	{
		this.application = application;
	}

	/// <summary>
	/// Retrieves the Outlook profile information, including accounts and
	/// stores.
	/// </summary>
	/// <returns>The Outlook profile information.</returns>
	public OutlookProfile GetProfileInfo()
	{
		NameSpace? session = application?.Session;

		return new OutlookProfile
		{
			Accounts = GetAccounts(session),
			Stores = GetStores(session)
		};
	}

	private static Collection<OutlookAccountInfo> GetAccounts(
		NameSpace? session)
	{
		Collection<OutlookAccountInfo> results = [];

		if (session != null)
		{
			foreach (Account account in session.Accounts)
			{
				Store? deliveryStore = null;

				try
				{
					deliveryStore = account.DeliveryStore;

					OutlookAccountInfo accountInfo = new()
					{
						DisplayName = account.DisplayName,
						UserName = account.UserName,
						SmtpAddress = account.SmtpAddress,
						AccountType = MapAccountType(account.AccountType),

						DeliveryStoreName = deliveryStore?.DisplayName,
						DeliveryStorePath = deliveryStore?.FilePath
					};

					results.Add(accountInfo);
				}
				finally
				{
					if (deliveryStore != null)
					{
						Marshal.ReleaseComObject(deliveryStore);
					}

					Marshal.ReleaseComObject(account);
				}
			}
		}

		return results;
	}

	private static Collection<OutlookStoreInfo> GetStores(NameSpace? session)
	{
		Collection<OutlookStoreInfo> results = [];

		if (session != null)
		{
			foreach (Store? store in session.Stores)
			{
				if (store != null)
				{
					try
					{
						OutlookStoreInfo info = new()
						{
							DisplayName = store.DisplayName,
							FilePath = store.FilePath,
							StoreId = store.StoreID,
							IsDataFileStore = store.IsDataFileStore,
							IsOpen = store.IsOpen
						};

						results.Add(info);
					}
					finally
					{
						Marshal.ReleaseComObject(store);
					}
				}
			}
		}

		return results;
	}

	private static OutlookAccountType MapAccountType(OlAccountType type)
	{
		return type switch
		{
			OlAccountType.olExchange => OutlookAccountType.Exchange,
			OlAccountType.olHttp => OutlookAccountType.Http,
			OlAccountType.olImap => OutlookAccountType.Imap,
			OlAccountType.olPop3 => OutlookAccountType.Pop3,
			OlAccountType.olOtherAccount => OutlookAccountType.Other,
			_ => OutlookAccountType.Unknown
		};
	}
}
