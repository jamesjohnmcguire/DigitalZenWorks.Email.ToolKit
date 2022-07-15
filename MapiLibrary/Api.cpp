#include "pch.h"

#include "Session.h"
#include "Store.h"

namespace MapiLibrary
{
	API void MapiTest()
	{
		std::cout << "This is a test." << std::endl;

		Session* session = new Session();

		std::vector<std::shared_ptr<Store>> stores = session->GetStores();

		session->Close();
		session = nullptr;
	}
}
