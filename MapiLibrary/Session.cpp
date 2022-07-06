#include "pch.h"
#include <vector>

#include <MAPIUtil.h>

#include "Session.h"
#include "Store.h"

namespace MapiLibrary
{
	MAPIINIT_0 MAPIINIT =
	{
		MAPI_INIT_VERSION,
		MAPI_MULTITHREAD_NOTIFICATIONS
	};

	Session::Session()
	{
		HRESULT result = MAPIInitialize(&MAPIINIT);

		if (result == S_OK)
		{
			result = MAPILogonEx(
				0,
				nullptr,
				nullptr,
				logonFlags,
				&mapiSession
			);
		}
	}

	Session::~Session()
	{
		Close();
	}

	void Session::Close()
	{
		int size = stores.size();
		for (int index = 0; index < size; index++)
		{
			Store* store = stores[index];
			delete store;
		}

		stores.clear();

		if (mapiSession != nullptr)
		{
			mapiSession->Release();
			mapiSession = nullptr;
		}

		MAPIUninitialize();
	}

	std::vector<Store*> Session::GetStores()
	{
		HRESULT result;
		LPMAPITABLE tableStores = nullptr;

		result = mapiSession->GetMsgStoresTable(MAPI_UNICODE, &tableStores);

		if (result == S_OK)
		{
			LPSRowSet rows = nullptr;

			result = HrQueryAllRows(tableStores,
				nullptr,
				nullptr,
				nullptr,
				0,
				&rows);

			if (result == S_OK)
			{
				ULONG entryIdLength;
				LPENTRYID entryId;
				LPSPropValue value;

				ULONG StoresCount = rows->cRows;

				for (ULONG index = 0; index < StoresCount; index++)
				{
					value = rows->aRow[index].lpProps;

					entryIdLength = value->Value.bin.cb;
					entryId = (LPENTRYID)value->Value.bin.lpb;

					Store* store = new Store(entryIdLength, entryId);
					stores.push_back(store);
				}
			}
		}

		return stores;
	}
}
