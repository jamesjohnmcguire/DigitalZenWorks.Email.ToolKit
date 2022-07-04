#include "pch.h"
#include "Store.h"

namespace MapiLibrary
{
	Store::Store(ULONG entryIdLengthIn, LPENTRYID entryIdIn)
		: entryIdLength(entryIdLengthIn), entryId(entryIdIn)
	{
	}

	Store::~Store()
	{
	}
}
