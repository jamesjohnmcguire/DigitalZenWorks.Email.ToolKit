#pragma once

#include "MapiProperties.h"

namespace MapiLibrary
{
	class Message
	{
		public:
			Message(LPMESSAGE messageIn);
			std::vector<byte> GetMessageHash();

		private:
			LPMESSAGE message;
	};
}
