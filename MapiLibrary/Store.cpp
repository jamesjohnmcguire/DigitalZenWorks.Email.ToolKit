#include "pch.h"
#include "Store.h"

namespace MapiLibrary
{
	Store::Store(
		LPMAPISESSION mapiSessionIn, ULONG entryIdLengthIn, LPENTRYID entryIdIn)
		:
			mapiSession(mapiSessionIn),
			entryIdLength(entryIdLengthIn),
			entryId(entryIdIn)
	{
	}

	Store::~Store()
	{
	}

	int Store::RemoveDuplicates()
	{
		int duplicatesRemoved = 0;

		HRESULT result = mapiSession->OpenMsgStore(
			0L,
			entryIdLength,
			entryId,
			nullptr,
			MAPI_BEST_ACCESS,
			&mapiDatabase);

		if (result == S_OK)
		{
			ULONG objectType = 0;

			result = mapiDatabase->OpenEntry(
				0,
				nullptr,
				nullptr,
				MAPI_MODIFY | MAPI_DEFERRED_ERRORS,
				&objectType,
				(LPUNKNOWN*)&rootFolder);
		}

		return duplicatesRemoved;
	}
}
