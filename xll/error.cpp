#pragma warning(disable: 4996)
#include <stdexcept>
#include "xll.h"
#include "registry.h"
#include "exports.h"
#include "error.h"

using namespace xll;

class reg_alert_level {
	Reg::Key key; 
	DWORD value;
public:
	// default to ERROR, WARNING, and INFO on
	reg_alert_level()
		: key(HKEY_CURRENT_USER, TEXT("Software\\KALX\\xll")), value(0x7)
	{
		try {
			value = key[TEXT("xll_alert_level")];
		}
		catch (...) {
			; // value gets set to default
		}
	}
	reg_alert_level& operator=(DWORD level)
	{
		key[TEXT("xll_alert_level")] = level;
        value = level;

		return *this;
	}
	operator DWORD() const
	{
		return value;
	}
} xll_alert_level;

DWORD XLL_ALERT_LEVEL(DWORD level)
{
    DWORD olevel = xll_alert_level;
    
    xll_alert_level = level;

    return olevel;
}

int 
XLL_ALERT(const char* text, const char* caption, DWORD level, UINT type, bool force)
{
	try {
		if ((xll_alert_level&level) || force) {
			if (IDCANCEL == MessageBoxA(GetForegroundWindow(), text, caption, MB_OKCANCEL|type))
				xll_alert_level = (xll_alert_level & ~level);
		}
	}
	catch (const std::exception& ex) {
		MessageBoxA(GetForegroundWindow(), ex.what(), "Alert", MB_OKCANCEL| MB_ICONERROR);
	}

	return static_cast<int>(xll_alert_level);
}

int 
XLL_ERROR(const char* e, bool force)
{
	return XLL_ALERT(e, "Error", XLL_ALERT_ERROR, MB_ICONERROR, force);
}
int 
XLL_WARNING(const char* e, bool force)
{
	return XLL_ALERT(e, "Warning", XLL_ALERT_WARNING, MB_ICONWARNING, force);
}
int 
XLL_INFO(const char* e, bool force)
{
	return XLL_ALERT(e, "Information", XLL_ALERT_INFO, MB_ICONINFORMATION, force);
}

XLL_CONST(WORD, XLL_ALERT_ERROR, XLL_ALERT_ERROR,
	"ALERT.LEVEL flag for errors[1].", "XLL", "")
XLL_CONST(WORD, XLL_ALERT_WARNING, XLL_ALERT_WARNING,
	"ALERT.LEVEL flag for warnings[2].", "XLL", "")
XLL_CONST(WORD, XLL_ALERT_INFO, XLL_ALERT_INFO,
	"ALERT.LEVEL flag for information[4].", "XLL", "")

AddIn xai_alert_level(
	Function(XLL_WORD, "xll_alert_level_", "XLL.ALERT.LEVEL")
	.Arguments({
		Arg(XLL_WORD, "level", "is the alert level mask to set.", "XLL_ALERT_ERROR()"),
	})
	.FunctionHelp("Set the current alert level using a mask and return the old mask.")
	.Category("XLL")
	.Documentation(R"(
The xll library can report errors, warnings, and information using pop-up alerts.
These can be turned on or off using the XLL_ALERT_* flags.
The value is stored in the registry to persist across Excel sessions.
)")
);
DWORD WINAPI xll_alert_level_(DWORD w)
{
#pragma XLLEXPORT
	DWORD oal = xll_alert_level;

	xll_alert_level = w;

	return oal;
}

#ifdef _DEBUG

struct test_dword {
	Reg::Key key;
	test_dword()
		: key(HKEY_CURRENT_USER, TEXT("tmp\\key"))
	{
		key[TEXT("foo")] = 123;
		DWORD dw;
		dw = key[TEXT("foo")];
		if (dw != 123) {
			MessageBox(0, L"dword failed", L"Error", MB_OK);
		}
	}
	~test_dword()
	{
		RegDeleteKey(key, TEXT("tmp\\key"));
	}
};
test_dword test_dword_{};

#endif // _DEBUG