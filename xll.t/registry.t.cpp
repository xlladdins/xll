// registry.t.cpp - test registry
#include <cassert>
#include "../xll/registry.h"

using namespace Reg;

int test_registry()
{
	{
		CreateKey key(HKEY_CURRENT_USER, TEXT("Software"));
		CreateKey::KeyIterator ki(key);
		PCTSTR s;
		s = *ki;
		++ki;
		s = *ki;
	}

	return 0;
}
int test_registry_ = test_registry();
