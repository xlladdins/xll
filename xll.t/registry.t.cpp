// registry.t.cpp - test registry
#include <cassert>
#include "../xll/registry.h"

using namespace Reg;

int test_registry()
{
	try {
		{
			Reg::SZ s(TEXT("abc"));
			chop(s, TEXT('b'));
			assert(s == TEXT("a"));
			chop(s, TEXT('b'));
			assert(s == TEXT("a"));
		}
		{
			Reg::SZ s(TEXT("abc"));
			tack(s, TEXT('d'));
			assert(s == TEXT("abcd"));
			tack(s, TEXT('d'));
			assert(s == TEXT("abcd"));
		}
		{
			Key hkcu(HKEY_CURRENT_USER, TEXT("Software"));
			assert(hkcu.disposition() == REG_OPENED_EXISTING_KEY);
			Key hkcc(HKEY_CURRENT_CONFIG, TEXT("Software"), KEY::READ);
			assert(hkcu != hkcc);
			Key hkcu2(HKEY_CURRENT_USER, TEXT("System"));
			assert(hkcu != hkcu2);
		}
		{
			Key key(HKEY_CURRENT_USER, TEXT("Microsoft"));
			Key::KeyIterator ki(key);
			assert(ki == ki);
			assert(!(ki != ki));
			Key::KeyIterator ke(key, true);
			assert(ki != ke);
			PCTSTR s;
			s = *ki;
			int icount = 0;
			for (auto i : key) {
				s = i;
				++icount;
			}

			Key key2(HKEY_CURRENT_USER, TEXT("Microsoft"));
			Key::KeyIterator ki2(key2);
			int ecount = 0;
			while (ki2) {
				++ki2;
				++ecount;
			}
			assert(icount == ecount);
		}
		{
			Key key(HKEY_CURRENT_USER, TEXT("Console"));
			Key::ValueIterator vi(key);
			while (vi) {
				if (vi.Type() == REG_DWORD) {
					DWORD dw;
					dw = *(DWORD*)vi.data();
				}
				if (vi.Type() == REG_SZ) {
					//PCTSTR s;
					//s = vi.data()
				}
				++vi;
			}
		}
	}
	catch (const std::exception& ex) {
		MessageBoxA(GetActiveWindow(), ex.what(), "Error", MB_OK);
	}

	return 0;
}
int test_registry_ = test_registry();
