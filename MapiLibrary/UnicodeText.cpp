#include "pch.h"
#include "UnicodeText.h"

namespace MapiLibrary
{
	std::string UnicodeText::GetUtf8Text(const std::wstring& wideString)
	{
		std::string utf8Text;

		if (!wideString.empty())
		{
			int sizeNeeded = WideCharToMultiByte(
				CP_UTF8,
				WC_ERR_INVALID_CHARS,
				&wideString[0],
				(int)wideString.size(),
				nullptr,
				0,
				nullptr,
				nullptr);

			utf8Text = std::string(sizeNeeded, 0);

			int result = WideCharToMultiByte(
				CP_UTF8,
				WC_ERR_INVALID_CHARS,
				&wideString[0],
				(int)wideString.size(),
				&utf8Text[0],
				sizeNeeded,
				nullptr,
				nullptr);
		}

		return utf8Text;
	}

	std::wstring UnicodeText::GetWideText(const std::string& utf8Text)
	{
		std::wstring wideText;

		if (!utf8Text.empty())
		{
			int sizeNeeded = MultiByteToWideChar(
				CP_UTF8,
				MB_ERR_INVALID_CHARS,
				&utf8Text[0],
				(int)utf8Text.size(),
				nullptr,
				0);

			wideText = std::wstring(sizeNeeded, 0);

			int result = MultiByteToWideChar(
				CP_UTF8,
				MB_ERR_INVALID_CHARS,
				&utf8Text[0],
				(int)utf8Text.size(),
				&wideText[0],
				sizeNeeded);
		}

		return wideText;
	}
}
