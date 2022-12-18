#include "pch.h"

#include "Log.h"
#include "Session.h"
#include "Store.h"

namespace MapiLibrary
{
	API void MapiTest()
	{
		Log log = Log();
		std::shared_ptr<Log> logPointer = std::make_shared<Log>(log);

		log.info("MapiTest Starting ");

		Session* session = new Session(logPointer);

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