#include "pch.h"
#include "Session.h"

namespace MapiLibrary
{
	Session::Session()
	{
		HRESULT result = MAPIInitialize(&MAPIINIT);

		if (result == S_OK)
		{
			result = MAPILogonEx(
				0,
				nullptr,
				nullptr,
				logonFlags,
				&mapiSession
			);
		}
	}
}
