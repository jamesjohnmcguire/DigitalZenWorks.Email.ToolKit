#include "pch.h"

#include "Folder.h"
#include "Store.h"
#include "UnicodeText.h"

namespace MapiLibrary
{
	Store::Store(
		LPMAPISESSION mapiSessionIn,
		ULONG entryIdLengthIn,
		LPENTRYID entryIdIn)
		:
			mapiSession(mapiSessionIn),
			entryIdLength(entryIdLengthIn),
			entryId(entryIdIn)
	{
	}

	Store::Store(
		LPMAPISESSION mapiSessionIn,
		ULONG entryIdLengthIn,
		LPENTRYID entryIdIn,
		std::string applicationName)
		: Store(mapiSessionIn, entryIdLengthIn, entryIdIn)
	{
		this->applicationName = applicationName;

		logger = spdlog::get(applicationName);
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
			unsigned long objectType = 0;

			LPSPropValue property = nullptr;
			result = HrGetOneProp(mapiDatabase, PR_DISPLAY_NAME, &property);
			const std::wstring storeName(property->Value.lpszW);

			std::string message =
				"Store: " + UnicodeText::GetUtf8Text(storeName);
			logger->info(message);

			result = mapiDatabase->OpenEntry(
				0,
				nullptr,
				nullptr,
				MAPI_MODIFY | MAPI_DEFERRED_ERRORS,
				&objectType,
				(LPUNKNOWN*)&rootFolder);

			if (result == S_OK && rootFolder != nullptr)
			{
				std::unique_ptr<Folder> folder =
					std::make_unique<Folder>(rootFolder, applicationName);

				duplicatesRemoved += folder->RemoveDuplicates();
			}

			if (rootFolder != nullptr)
			{
				rootFolder->Release();
				rootFolder = nullptr;
			}
		}

		return duplicatesRemoved;
	}
}
