#include "pch.h"

#include "MapiProperties.h"
#include "Folder.h"
#include <EdkMdb.h>

namespace MapiLibrary
{
	Folder::Folder(LPMAPIFOLDER mapiFolderIn)
		: mapiFolder(mapiFolderIn)
	{

	}

	int Folder::RemoveDuplicates()
	{
		int duplicatesRemoved = 0;

		LPSPropValue property = nullptr;
		HRESULT result = HrGetOneProp(mapiFolder, PR_DISPLAY_NAME, &property);
		LPWSTR name = property->Value.lpszW;
		std::wcout << "Folder: " << name << std::endl;

		std::vector<std::shared_ptr<Folder>> folders = GetChildFolders();

		size_t size = folders.size();
		for (size_t index = 0; index < size; index++)
		{
			std::shared_ptr<Folder> folderPointer = folders[index];
			Folder* folder = folderPointer.get();
			duplicatesRemoved += folder->RemoveDuplicates();
		}

		duplicatesRemoved += RemoveDuplicatesInThisFolder();

		return duplicatesRemoved;
	}

	std::vector<std::shared_ptr<Folder>> Folder::GetChildFolders()
	{
		std::vector<std::shared_ptr<Folder>> folders;

		LPMAPITABLE childTable;

		HRESULT result = mapiFolder->GetHierarchyTable(0, &childTable);

		if (result == S_OK)
		{
			SizedSPropTagArray(2, folderTags) =
			{
				2,
				{
					PR_DISPLAY_NAME,
					PR_ENTRYID
				}
			};

			result = childTable->SetColumns((LPSPropTagArray)&folderTags, 0);

			if (result == S_OK)
			{
				unsigned long rowCount = 0;
				result = childTable->GetRowCount(0, &rowCount);

				if (result == S_OK && rowCount > 0)
				{
					long rowsSeeked = 0;
					result = childTable->SeekRow(
						BOOKMARK_BEGINNING,
						0,
						nullptr);

					if (result == S_OK)
					{
						folders = QueryForChildFolders(childTable, rowCount);
					}
				}
			}

			childTable->Release();
			childTable = nullptr;
		}

		return folders;
	}

	std::shared_ptr<Folder> Folder::GetChildFolder(SRow row)
	{
		std::shared_ptr<Folder> folder = nullptr;

		LPSPropValue properties = row.lpProps;
		SPropValue property1 = properties[1];
		unsigned long tag = property1.ulPropTag;

		if (tag == PR_ENTRYID)
		{
			LPMAPIFOLDER childFolder = nullptr;
			unsigned long objectType = 0;

			unsigned long childEntryIdSize = property1.Value.bin.cb;
			LPENTRYID childEntryId = (LPENTRYID)property1.Value.bin.lpb;

			HRESULT result = mapiFolder->OpenEntry(
				childEntryIdSize,
				childEntryId,
				nullptr,
				MAPI_MODIFY,
				&objectType,
				(IUnknown**)&childFolder);

			if (result == S_OK)
			{
				folder = std::make_shared<Folder>(childFolder);
			}
		}

		return folder;
	}

	std::vector<std::shared_ptr<Folder>> Folder::QueryForChildFolders(
		LPMAPITABLE childTable, unsigned long rowCount)
	{
		std::vector<std::shared_ptr<Folder>> folders;

		LPSRowSet rows = nullptr;
		HRESULT result = childTable->QueryRows(rowCount, 0, &rows);

		if (result == S_OK)
		{
			int rowCount = rows->cRows;

			for (int index = 0; index < rowCount; index++)
			{
				SRow row = rows->aRow[index];
				std::shared_ptr<Folder> folder =
					GetChildFolder(row);

				if (folder != nullptr)
				{
					folders.push_back(folder);
				}
			}

			FreeProws(rows);
			rows = nullptr;
		}

		return folders;
	}

	int Folder::RemoveDuplicatesInThisFolder()
	{
		int duplicatesRemoved = 0;

		LPMAPITABLE mapiTable = nullptr;

		HRESULT result = mapiFolder->GetContentsTable(0, &mapiTable);

		if (result == S_OK)
		{
			SizedSPropTagArray(4, itemTags) =
			{
				3,
				{
					PR_ENTRYID,
					PR_MESSAGE_CLASS,
					PR_SUBJECT
				}
			};

			SRowSet* rows;
			result = HrQueryAllRows(
				mapiTable, (SPropTagArray*)&itemTags, NULL, NULL, 0, &rows);
			if (result == S_OK)
			{
				unsigned long count = rows->cRows;

				for (unsigned long index = 0; index < count; index++)
				{
					SRow row = rows->aRow[index];
					LPSPropValue properties = row.lpProps;
					SPropValue property = properties[0];

					unsigned long tag = property.ulPropTag;

					if (tag == PR_ENTRYID)
					{
						LPMAPIFOLDER childFolder = nullptr;
						unsigned long objectType = 0;

						unsigned long entryIdSize = property.Value.bin.cb;
						LPENTRYID entryId = (LPENTRYID)property.Value.bin.lpb;

						LPMESSAGE message;
						ULONG messageType;
						result = mapiFolder->OpenEntry(
							entryIdSize,
							entryId,
							nullptr,
							0,
							&messageType,
							(IUnknown**)&message);

						SizedSPropTagArray(53, itemTags) =
						{
							53,
							{
								PR_ACCESS,
								PR_ACCESS_LEVEL,
								PR_BODY,
								PR_CLIENT_SUBMIT_TIME,
								PR_CONVERSATION_INDEX,
								PR_CREATION_TIME,
								PR_DISPLAY_NAME,
								PR_SUBJECT,
								PR_SENT_REPRESENTING_NAME,
								PR_MESSAGE_DELIVERY_TIME,
								PR_DISPLAY_BCC,
								PR_DISPLAY_CC,
								PR_DISPLAY_TO,
								PR_HASATTACH,
								PR_HTML,
								PR_IMPORTANCE,
								PR_INTERNET_CPID,
								PR_LAST_MODIFICATION_TIME,
								PR_MAPPING_SIGNATURE,
								PR_MDB_PROVIDER,
								PR_MESSAGE_ATTACHMENTS,
								PR_MESSAGE_CLASS,
								PR_MESSAGE_DELIVERY_TIME,
								PR_MESSAGE_FLAGS,
								PR_MESSAGE_RECIPIENTS,
								PR_NORMALIZED_SUBJECT,
								PR_OBJECT_TYPE,
								PR_RECORD_KEY,
								PR_RTF_COMPRESSED,
								PR_RTF_IN_SYNC,
								PR_RECEIVED_BY_ADDRTYPE,
								PR_RECEIVED_BY_EMAIL_ADDRESS,
								PR_RECEIVED_BY_ENTRYID,
								PR_RECEIVED_BY_NAME,
								PR_RECEIVED_BY_SEARCH_KEY,
								PR_REPLY_RECIPIENT_ENTRIES,
								PR_REPLY_RECIPIENT_NAMES,
								PR_SEARCH_KEY,
								PR_SENDER_ADDRTYPE,
								PR_SENDER_EMAIL_ADDRESS,
								PR_SENDER_NAME,
								PR_SENT_REPRESENTING_ADDRTYPE,
								PR_SENT_REPRESENTING_EMAIL_ADDRESS,
								PR_SENT_REPRESENTING_NAME,
								PR_SUBJECT_PREFIX,
								PR_SUBJECT,
								PR_INTERNET_MESSAGE_ID,
								PR_SENDER_ENTRYID,
								PR_SENDER_SEARCH_KEY,
								PR_SENT_REPRESENTING_ENTRYID,
								PR_SENT_REPRESENTING_NAME,
								PR_SENT_REPRESENTING_SEARCH_KEY,
								PR_TRANSPORT_MESSAGE_HEADERS,
							}
						};
					}
				}
			}
		}

		return duplicatesRemoved;
	}
}
