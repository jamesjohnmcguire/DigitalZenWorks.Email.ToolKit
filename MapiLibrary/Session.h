#pragma once

#include "MapiLibrary.h"

namespace MapiLibrary
{
	class Session
	{
		public:
			DllExport Session();
			~Session() = default;

		private:
			LPMAPISESSION mapiSession;
	};
}
