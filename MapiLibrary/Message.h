#pragma once

#include "MapiProperties.h"

namespace MapiLibrary
{
	class Message
	{
		public:
			Message(LPMESSAGE messageIn);
			std::string GetMessageHash();

		private:
			LPMESSAGE message;
	};
}
