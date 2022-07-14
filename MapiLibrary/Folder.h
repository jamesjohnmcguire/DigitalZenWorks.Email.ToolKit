#pragma once

namespace MapiLibrary
{
	class Folder
	{
		public:
			Folder(LPMAPIFOLDER folder);
			int RemoveDuplicates();
	
		private:
			LPMAPIFOLDER mapiFolder;
	};
}
