#include "pch.h"

#include "Log.h"
#include "Session.h"
#include "Store.h"

namespace MapiLibrary
{
	static const std::string logger_name = "NC";

	API(void) MapiTest()
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

	API(void) test(std::string message)
	{
		auto logger = spdlog::get(logger_name);
		if (logger)
		{
			logger->debug("{}::{}", __FUNCTION__, message);
		}
	}
}
