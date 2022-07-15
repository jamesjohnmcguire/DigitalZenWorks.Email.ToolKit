#include "pch.h"

#include "Folder.h"
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
			SPropValue* ipmEntryId;
			result = HrGetOneProp(
				mapiDatabase, PR_IPM_SUBTREE_ENTRYID, &ipmEntryId);

			if (result == S_OK)
			{
				ULONG objectType = 0;
				ULONG propSize = UlPropSize(ipmEntryId);

				result = mapiDatabase->OpenEntry(
					propSize,
					(LPENTRYID)ipmEntryId,
					nullptr,
					MAPI_MODIFY,
					&objectType,
					(LPUNKNOWN*)&rootFolder);

				//result = mapiDatabase->OpenEntry(
				//	0,
				//	nullptr,
				//	nullptr,
				//	MAPI_MODIFY | MAPI_DEFERRED_ERRORS,
				//	&objectType,
				//	(LPUNKNOWN*)&rootFolder);

				std::unique_ptr<Folder> folder =
					std::make_unique<Folder>(rootFolder);

				duplicatesRemoved += folder->RemoveDuplicates();
			}
		}

		return duplicatesRemoved;
	}
}
