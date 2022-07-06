#include "pch.h"

#include <iostream>

#include "Session.h"
#include "Store.h"

namespace MapiLibrary
{
	LIB_API(void) MapiTest()
	{
		std::cout << "This is a test." << std::endl;

		Session* session = new Session();

		std::vector<Store*> stores = session->GetStores();

		session->Close();
		session = nullptr;
	}
}
