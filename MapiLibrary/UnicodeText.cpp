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

	char* UnicodeText::GetUtf8Text(const wchar_t* wideString)
	{
		char* utf8Text = nullptr;

		if (wideString != nullptr)
		{
			int sizeNeeded = WideCharToMultiByte(
				CP_UTF8,
				WC_ERR_INVALID_CHARS,
				wideString,
				-1,
				nullptr,
				0,
				nullptr,
				nullptr);

			utf8Text = (char*)malloc(sizeNeeded);

			int result = WideCharToMultiByte(
				CP_UTF8,
				WC_ERR_INVALID_CHARS,
				wideString,
				-1,
				utf8Text,
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

	void UnicodeText::SetConsole()
	{
		CONSOLE_FONT_INFOEX cfi;
		cfi.cbSize = sizeof cfi;
		cfi.nFont = 0;
		cfi.dwFontSize.X = 10;
		cfi.dwFontSize.Y = 20;
		cfi.FontFamily = FF_DONTCARE;
		cfi.FontWeight = FW_NORMAL;

		wcscpy_s(cfi.FaceName, 20, L"MS Mincho");

		HANDLE standardHandle = GetStdHandle(STD_OUTPUT_HANDLE);
		SetCurrentConsoleFontEx(standardHandle, FALSE, &cfi);

		SetConsoleOutputCP(CP_UTF8);
	}
}
