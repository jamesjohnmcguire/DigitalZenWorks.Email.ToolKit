#pragma once

#include "MapiLibrary.h"

namespace MapiLibrary
{
	MAPIINIT_0 MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };

	class Session
	{
		public:
			DllExport Session();
			~Session() = default;

		private:
			ULONG logonFlags = MAPI_ALLOW_OTHERS | MAPI_EXTENDED |
				MAPI_NO_MAIL | MAPI_USE_DEFAULT | MAPI_UNICODE;
			LPMAPISESSION mapiSession{};
	};
}
