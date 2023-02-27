#pragma once

#include "MapiLibrary.h"

namespace MapiLibrary
{
	class UnicodeText
	{
		public:
			DllExport static char* GetUtf8Text(const wchar_t* wideString);
			DllExport static std::string GetUtf8Text(
				const std::wstring& wideString);
			DllExport static std::wstring GetWideText(
				const std::string& utf8Text);

			DllExport static void SetConsole();
	};
}
