#pragma once

#include "Log.h"
#include "MapiProperties.h"

namespace MapiLibrary
{
	class Message
	{
		public:
			Message(LPMESSAGE messageIn);
			Message(LPMESSAGE messageIn, std::string applicationName);
			std::string GetMessageHash();

		private:
			std::string applicationName;
			std::shared_ptr<spdlog::logger> logger = nullptr;
			LPMESSAGE message;
	};
}
