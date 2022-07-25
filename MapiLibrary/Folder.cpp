#include "pch.h"

#include "Folder.h"

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
			SizedSPropTagArray(5, itemTags) =
			{
				5,
				{
					PR_ENTRYID,
					PR_DISPLAY_NAME,
					PR_SUBJECT,
					PR_SENT_REPRESENTING_NAME,
					PR_MESSAGE_DELIVERY_TIME
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

				}
			}
		}

		return duplicatesRemoved;
	}
}
