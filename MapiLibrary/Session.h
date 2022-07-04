#pragma once
#include <vector>

#include "MapiLibrary.h"
#include "Store.h"

namespace MapiLibrary
{
	MAPIINIT_0 MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };

	class Session
	{
		public:
			DllExport Session();
			DllExport ~Session();

			DllExport void Close();
			DllExport std::vector<Store*> GetStores();

		private:
			ULONG logonFlags = MAPI_ALLOW_OTHERS | MAPI_EXTENDED |
				MAPI_NO_MAIL | MAPI_USE_DEFAULT | MAPI_UNICODE;
			LPMAPISESSION mapiSession{};
			std::vector<Store*> stores;
	};
}
