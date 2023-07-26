// async.cpp - Asynchonous callback
#include <thread>
#include <chrono>
#include <coroutine>
#include "../xll/xll.h"

using namespace xll;

struct ARGS {
	OPER12 arg, handle;
};

int my_async(WORD ms, OPER handle)
{
	std::this_thread::sleep_for(std::chrono::milliseconds(ms));
	OPER arg;
	arg = ms;

	int ret = ::Excel12(xlAsyncReturn, 0, 2, &handle, &arg);
	
	return ret;
}

AddIn xai_async(
	Function(XLL_VOID, "xll_async", "XLL.ASYNC")
	.Arguments({
		Arg(XLL_WORD, "ms", "is the argument")
		})
	.Asynchronous()
	//.ThreadSafe()
);
void WINAPI xll_async(WORD ms, LPOPER12 handle)
{
#pragma XLLEXPORT
	{
		std::jthread t(my_async, ms, *handle);
		t.detach();
	}
	
	return;
}