#pragma once
#include "DeclarationsSpecifications.h"
#include "Log.h"

namespace MapiLibrary
{
	class Store
	{
		public:
			DllExport Store(
				LPMAPISESSION mapiSession, ULONG entryIdLength, LPENTRYID entryId);
			DllExport Store(
				LPMAPISESSION mapiSession,
				ULONG entryIdLength,
				LPENTRYID entryId,
				std::shared_ptr<Log> logger);
			DllExport ~Store();
			DllExport int RemoveDuplicates();

		private:
			LPENTRYID entryId;
			ULONG entryIdLength;
			std::shared_ptr<Log> logger = nullptr;
			LPMDB mapiDatabase{};
			LPMAPISESSION mapiSession;
			LPMAPIFOLDER rootFolder{};
	};
}
