#pragma once

#include "DeclarationsSpecifications.h"
#include "Log.h"
#include "Store.h"

namespace MapiLibrary
{
	class Session
	{
		public:
			DllExport Session();
			DllExport Session(std::shared_ptr<Log> logger);
			DllExport ~Session();

			DllExport void Close();
			DllExport std::vector<std::shared_ptr<Store>> GetStores();

		private:
			ULONG logonFlags = MAPI_ALLOW_OTHERS | MAPI_EXTENDED |
				MAPI_NO_MAIL | MAPI_USE_DEFAULT | MAPI_UNICODE;
			LPMAPISESSION mapiSession{};

			std::shared_ptr<Log> logger = nullptr;
			std::vector<std::shared_ptr<Store>> stores;
	};
}