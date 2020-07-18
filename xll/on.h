// on.h - xlcOnXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "auto.h"
#include "excel.h"

// used with On<Key>
#define ON_SHIFT       L"+"
#define ON_CTRL        L"^"
#define ON_ALT         L"%"
#define ON_COMMAND     L"*"
#define ON_ENTER       L"~"
#define ON_BACKSPACE   L"{BACKSPACE}" 
#define ON_BS          L"{BS}" 
#define ON_BREAK       L"{BREAK}" 
#define ON_CAPS_LOCK   L"{CAPSLOCK}" 
#define ON_CLEAR       L"{CLEAR}" 
#define ON_DELETE      L"{DELETE}" 
#define ON_DEL         L"{DEL}" 
#define ON_DOWN_ARROW  L"{DOWN}" 
#define ON_END         L"{END}" 
#define ON_ENTER_NUM   L"{ENTER}" 
#define ON_ESCAPE      L"{ESCAPE}" 
#define ON_ESC         L"{ESC}" 
#define ON_F1          L"{F1}" 
#define ON_F2          L"{F2}" 
#define ON_F3          L"{F3}" 
#define ON_F4          L"{F4}" 
#define ON_F5          L"{F5}" 
#define ON_F6          L"{F6}" 
#define ON_F7          L"{F7}" 
#define ON_F8          L"{F8}" 
#define ON_F9          L"{F9}" 
#define ON_F10         L"{F10}" 
#define ON_F11         L"{F11}" 
#define ON_F12         L"{F12}" 
#define ON_F13         L"{F13}" 
#define ON_F14         L"{F14}" 
#define ON_F15         L"{F15}" 
#define ON_HELP        L"{HELP}" 
#define ON_HOME        L"{HOME}" 
#define ON_INSERT      L"{INSERT}" 
#define ON_LEFT_ARROW  L"{LEFT}" 
#define ON_NUM_LOCK    L"{NUMLOCK}" 
#define ON_PAGE_DOWN   L"{PGDN}" 
#define ON_PAGE_UP     L"{PGUP}" 
#define ON_RETURN      L"{RETURN}" 
#define ON_RIGHT_ARROW L"{RIGHT}" 
#define ON_SCROLL_LOCK L"{SCROLLLOCK}" 
#define ON_TAB         L"{TAB}" 
#define ON_UP_ARROW    L"{UP}" 

namespace xll {

	struct Data {
		static const int On = xlcOnData;
	};
	struct Doubleclick {
		static const int On = xlcOnDoubleclick;
	};
	struct Entry {
		static const int On = xlcOnEntry;
	};
	struct Key {
		static const int On = xlcOnKey;
	};
	struct Recalc {
		static const int On = xlcOnRecalc;
	};
	struct Sheet {
		static const int On = xlcOnSheet;
	};
	struct Time {
		static const int On = xlcOnTime;
	};
	struct Window {
		static const int On = xlcOnWindow;
	};
	
	// static On<Key> ok("shortcut", "MACRO");
	template<class Key>
	class On {
		using xcstr = const XCHAR*;
	public:
		On(xcstr text, xcstr macro)
		{
			Auto<Open> xao([text, macro]() {
				return ExcelX(Key::On, OPERX(text), OPERX(macro)) == true;
			});
			Auto<Close> xac([text]() {
				return ExcelX(Key::On, OPERX(text)) == true;
			});
		}
		On(xcstr text, xcstr macro, bool activate)
		{
			Auto<Open> xao([text, macro, activate]() {
				return ExcelX(Key::On, OPERX(text), OPERX(macro), OPERX(activate)) == true;
			});
			Auto<Close> xac([text]() {
				return ExcelX(Key::On, OPERX(text)) == true;
			});
		}
		On(const OPERX& time, xcstr macro, const OPERX& tolerance, bool insert)
		{
			Auto<Open> xao([time, macro, tolerance, insert]() {
				return ExcelX(Key::On, OPERX(time), OPERX(macro), OPERX(tolerance), OPERX(insert)) == true;
			});
			Auto<Close> xac([]() {
				return ExcelX(Key::On) == true;
			});
		}
	};
} // namespace xll