#include "pch.h"

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
		if (logger == nullptr)
		{
			logger = spdlog::stdout_color_mt("console");
			logger->set_pattern("%+");
			logger->set_level(spdlog::level::trace);
		}

		logger->info("Starting Session");
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

	Session::Session(std::shared_ptr<spdlog::logger> logger) : Session()
	{
		this->logger = logger;
	}

	Session::~Session()
	{
		Close();
	}

	void Session::Close()
	{
		stores.clear();

		if (mapiSession != nullptr)
		{
			mapiSession->Logoff(0, 0, 0);
			mapiSession->Release();
			mapiSession = nullptr;
		}

		MAPIUninitialize();
	}

	std::vector<std::shared_ptr<Store>> Session::GetStores()
	{
		HRESULT result;
		LPMAPITABLE tableStores = nullptr;

		result = mapiSession->GetMsgStoresTable(0, &tableStores);

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
				unsigned long entryIdLength;
				LPENTRYID entryId;
				LPSPropValue value;

				unsigned long StoresCount = rows->cRows;

				for (unsigned long index = 0; index < StoresCount; index++)
				{
					value = rows->aRow[index].lpProps;

					entryIdLength = value->Value.bin.cb;
					entryId = (LPENTRYID)value->Value.bin.lpb;

					std::shared_ptr<Store> store = std::make_shared<Store>(
						mapiSession, entryIdLength, entryId);

					stores.push_back(store);
				}
			}
		}

		return stores;
	}
}
