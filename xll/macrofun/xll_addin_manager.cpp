// addin_manager.cpp
#include "xll/xll/splitpath.h"
#include "addin_manager.h"

using namespace xll;

AddIn xai_aim_add(Macro("xll_aim_add", "AIM.ADD"));
int WINAPI xll_aim_add(void)
{
#pragma XLLEXPORT
	try {
		AddinManager::Add(OPER("xllmonte"));
	}
	catch (...) {
		// ignore
	}

	return true;
}

AddIn xai_aim_remove(Macro("xll_aim_remove", "AIM.REMOVE"));
int WINAPI xll_aim_remove(void)
{
#pragma XLLEXPORT
	try {
		AddinManager::Remove(OPER("xllmonte"));
	}
	catch (...) {
		// ignore
	}

	return true;
}

AddIn xai_aim_new(Macro("xll_aim_new", "AIM.NEW"));
int WINAPI xll_aim_new(void)
{
#pragma XLLEXPORT
	try {
		AddinManager::New();
	}
	catch (...) {
		// ignore
	}

	return true;
}

AddIn xai_aim_delete(Macro("xll_aim_delete", "AIM.DELETE"));
int WINAPI xll_aim_delete(void)
{
#pragma XLLEXPORT
	try {
		AddinManager::Delete();
	}
	catch (...) {
		// ignore
	}

	return true;
}