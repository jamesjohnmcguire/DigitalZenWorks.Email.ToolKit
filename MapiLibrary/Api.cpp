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

		size_t count = stores.size();
		for (size_t index = 0; index < count; index++)
		{
			std::shared_ptr<Store> storePointer = stores[index];
			Store* store = storePointer.get();

			store->RemoveDuplicates();
		}

		session->Close();
		session = nullptr;
	}
}
