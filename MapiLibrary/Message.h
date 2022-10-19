#pragma once

#include "Log.h"
#include "MapiProperties.h"

namespace MapiLibrary
{
	class Message
	{
		public:
			Message(LPMESSAGE messageIn);
			Message(LPMESSAGE messageIn, std::shared_ptr<Log> logger);
			std::string GetMessageHash();

		private:
			std::shared_ptr<Log> logger = nullptr;
			LPMESSAGE message;
	};
}
