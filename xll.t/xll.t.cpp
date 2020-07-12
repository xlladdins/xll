// xll.t.cpp - test xll add-in library
#include <cassert>

extern "C" int __declspec(dllexport) __stdcall xlAutoOpen(void);

int main()
{
	//assert (xlAutoOpen());

	return 0;
}