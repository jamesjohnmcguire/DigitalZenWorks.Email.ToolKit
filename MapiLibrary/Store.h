#pragma once
#include "DeclarationsSpecifications.h"

namespace MapiLibrary
{
	class Store
	{
	public:
		DllExport Store(
			LPMAPISESSION mapiSession, ULONG entryIdLength, LPENTRYID entryId);
		DllExport ~Store();
		DllExport int RemoveDuplicates();

	private:
			LPENTRYID entryId;
			ULONG entryIdLength;
			LPMDB mapiDatabase{};
			LPMAPISESSION mapiSession;
			LPMAPIFOLDER rootFolder{};
	};
}
