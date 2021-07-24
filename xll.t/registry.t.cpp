// registry.t.cpp - test registry
#include <cassert>
#include "../xll/registry.h"

using namespace Reg;

int test_registry()
{
	try {
		using SZ = std::basic_string<TCHAR>;
		{
			SZ s(TEXT("abc"));
			chop(s, TEXT('b'));
			assert(s == TEXT("a"));
			chop(s, TEXT('b'));
			assert(s == TEXT("a"));
		}
		{
			SZ s(TEXT("abc"));
			tack(s, TEXT('d'));
			assert(s == TEXT("abcd"));
			tack(s, TEXT('d'));
			assert(s == TEXT("abcd"));
		}
		{
			Value v;
			assert(v.name == _T(""));
			assert(v.index == -1);
			assert(v.type == REG_NONE);
			assert(v.data.size() == 0);

			assert(v == v);
			assert(!(v != v));

			Value v2(v);
			v = v2;
			assert(v == v2);
		}
		{
			Value v(_T("foo"));
			assert(v.name == _T("foo"));
			assert(v.type == REG_NONE);
		}
		{
			Value v(_T("foo"));
			v = _T("bar");
			assert(v.name == _T("foo"));
			assert(v.type == REG_SZ);
			PCTSTR s = v;
			assert(0 == _tcscmp(s, _T("bar")));
		}
		{
			Value v(_T("foo"));
			v = 123;
			assert(v.type == REG_DWORD);
			DWORD dw;
			dw = v;
			assert(v == 123);
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
			Key::KeyIterator ke(key);
			assert(ki == ke);
			PCTSTR s;
			s = *ki;
			int icount = 0;
			for (auto i : key.Keys()) {
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
			assert(icount > 0);
		}
		{
			Key key(HKEY_CURRENT_USER, TEXT("Console"));
			Key::ValueIterator vi(key);
			int icount = 0;
			while (vi) {
				auto v = *vi;
				if (v.type == REG_DWORD) {
					++icount;
				}
				++vi;
			}
			int ecount = 0;
			for (const auto& v : key.Values()) {
				if (v.type == REG_DWORD) {
					++ecount;
				}
			}
			assert(icount == ecount);
			assert(icount > 0);
		}
		{
			Key key(HKEY_CURRENT_USER, _T("foo"));
			key[_T("bar")] = 123;
			DWORD dw;
			dw = key[_T("bar")];
			assert(dw == 123);
			key[_T("baz")] = _T("blah");
			auto p = key[_T("baz")];
			PCTSTR ps = p;
			// ps = key[_T("baz")]; // calls Proxy destructor
			assert(0 == _tcscmp(ps, _T("blah")));
			std::basic_string<TCHAR> s(key[_T("baz")]);
			assert(s == _T("blah"));

			RegDeleteKeyValue(HKEY_CURRENT_USER, _T("foo"), _T("bar"));
			RegDeleteKeyValue(HKEY_CURRENT_USER, _T("foo"), _T("baz"));
			RegDeleteKey(HKEY_CURRENT_USER, _T("foo"));
		}
	}
	catch (const std::exception& ex) {
		MessageBoxA(GetActiveWindow(), ex.what(), "Error", MB_OK);
	}

	return 0;
}
int test_registry_ = test_registry();
