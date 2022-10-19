#pragma once

namespace MapiLibrary
{
	class UnicodeText
	{
		public:
			static std::string GetUtf8Text(const std::wstring& wideString);
			static std::wstring GetWideText(const std::string& utf8Text);
	};
}
