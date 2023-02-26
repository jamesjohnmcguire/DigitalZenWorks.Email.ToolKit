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
				std::string applicationName);
			DllExport ~Store();
			DllExport int RemoveDuplicates();

		private:
			std::string applicationName;
			LPENTRYID entryId;
			ULONG entryIdLength;
			std::shared_ptr<spdlog::logger> logger = nullptr;
			LPMDB mapiDatabase{};
			LPMAPISESSION mapiSession;
			LPMAPIFOLDER rootFolder{};
	};
}
