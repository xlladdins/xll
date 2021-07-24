// xll_install.cpp - install and uninstall add-ins
#include "xll_addin_manager.h"

using namespace xll;

AddIn xai_install(
	Macro("xll_install", "XLL.INSTALL")
	.Category("XLL")
	.Documentation(Installer::install_doc)
);
int WINAPI xll_install()
{
#pragma XLLEXPORT

	Installer install(Installer::Template());

	try {
		if (!Document::Sheet()) {
			Workbook::New();
		}
		install.Install(Excel(xlGetName));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what(), true);

		return FALSE;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return FALSE;
	}

	return install.State() == Installer::EXISTS + Installer::KNOWN;
}

AddIn xai_uninstall(Macro("xll_uninstall", "XLL.UNINSTALL"));
int WINAPI xll_uninstall()
{
#pragma XLLEXPORT
	Installer install(Installer::Template());

	try {
		xll::path path(Excel(xlGetName).as_cstr());
		if (path) {
			install.Uninstall(OPER(path.fname));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what(), true);

		return FALSE;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return FALSE;
	}

	return install.State() == Installer::INIT;
}
