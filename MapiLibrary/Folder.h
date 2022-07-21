#pragma once

namespace MapiLibrary
{
	class Folder
	{
		public:
			Folder(LPMAPIFOLDER folder);
			int RemoveDuplicates();
	
		private:
			std::shared_ptr<Folder> GetChildFolder(SRow row);
			std::vector<std::shared_ptr<Folder>> GetChildFolders();
			std::vector<std::shared_ptr<Folder>> QueryForChildFolders(
				LPMAPITABLE childTable, unsigned long rowCount);
			int RemoveDuplicatesInThisFolder();

			LPMAPIFOLDER mapiFolder;
	};
}
