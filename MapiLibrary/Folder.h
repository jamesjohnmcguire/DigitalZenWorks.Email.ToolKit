#pragma once

#include "Log.h"

namespace MapiLibrary
{
	class Folder
	{
		public:
			Folder(LPMAPIFOLDER folder);
			Folder(LPMAPIFOLDER folder, std::shared_ptr<Log> logger);
			int RemoveDuplicates();
	
		private:
			std::shared_ptr<Folder> GetChildFolder(SRow row);
			std::vector<std::shared_ptr<Folder>> GetChildFolders();
			std::vector<std::shared_ptr<Folder>> QueryForChildFolders(
				LPMAPITABLE childTable, unsigned long rowCount);
			int RemoveDuplicatesInThisFolder();

			std::shared_ptr<Log> logger = nullptr;
			LPMAPIFOLDER mapiFolder;
	};
}
