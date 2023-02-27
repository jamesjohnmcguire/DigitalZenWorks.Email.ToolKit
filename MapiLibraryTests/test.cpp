#include "pch.h"

#include <iostream>
#include <memory>
#include <vector>

#include "../MapiLibrary/Session.h"
#include "../MapiLibrary/Store.h"

using namespace MapiLibrary;

TEST(TestSanityCheck, SanityCheck)
{
	EXPECT_EQ(1, 1);
	EXPECT_TRUE(true);
}

TEST(TestGetMessage, GetMessage)
{
	Session* session = new Session("MapiLibrary.Tests");

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

	EXPECT_EQ(1, 1);
	EXPECT_TRUE(true);
}
