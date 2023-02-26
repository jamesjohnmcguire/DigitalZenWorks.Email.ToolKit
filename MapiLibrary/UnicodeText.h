#pragma once

#include "MapiLibrary.h"

namespace MapiLibrary
{
	class UnicodeText
	{
		public:
			static std::string GetUtf8Text(const std::wstring& wideString);
			static std::wstring GetWideText(const std::string& utf8Text);

			static char* GetUtf8Text(const wchar_t* wideString);
			DllExport static void SetConsole();
	};
}
