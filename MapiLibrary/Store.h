#pragma once
#include "DeclarationsSpecifications.h"

namespace MapiLibrary
{
	class Store
	{
	public:
		DllExport Store(ULONG entryIdLength, LPENTRYID entryId);
		DllExport ~Store();

	private:
			LPENTRYID entryId;
			ULONG entryIdLength;
	};
}
