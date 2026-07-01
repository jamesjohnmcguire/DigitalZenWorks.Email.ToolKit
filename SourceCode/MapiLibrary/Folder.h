#pragma once

#include "Log.h"

namespace MapiLibrary
{
	class Folder
	{
		public:
			Folder(LPMAPIFOLDER folder);
			Folder(LPMAPIFOLDER folder, std::string applicationName);
			int RemoveDuplicates();
	
		private:
			std::shared_ptr<Folder> GetChildFolder(SRow row);
			std::vector<std::shared_ptr<Folder>> GetChildFolders();
			std::vector<std::shared_ptr<Folder>> QueryForChildFolders(
				LPMAPITABLE childTable, unsigned long rowCount);
			int RemoveDuplicatesInThisFolder();

			std::string applicationName;
			std::shared_ptr<spdlog::logger> logger = nullptr;
			LPMAPIFOLDER mapiFolder;
	};
}
