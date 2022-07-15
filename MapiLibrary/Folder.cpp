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

		duplicatesRemoved += RemoveDuplicatesInThisFolder();

		return duplicatesRemoved;
	}

	int Folder::RemoveDuplicatesInThisFolder()
	{
		int duplicatesRemoved = 0;

		return duplicatesRemoved;
	}
}
