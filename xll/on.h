// on.h - xlcOnXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "auto.h"
#include "excel.h"

// For use with On<Key>
#define ON_SHIFT       "+"
#define ON_CTRL        "^"
#define ON_ALT         "%"
#define ON_COMMAND     "*"
#define ON_ENTER       "~"
#define ON_BACKSPACE   "{BACKSPACE}" 
#define ON_BS          "{BS}" 
#define ON_BREAK       "{BREAK}" 
#define ON_CAPS_LOCK   "{CAPSLOCK}" 
#define ON_CLEAR       "{CLEAR}" 
#define ON_DELETE      "{DELETE}" 
#define ON_DEL         "{DEL}" 
#define ON_DOWN_ARROW  "{DOWN}" 
#define ON_END         "{END}" 
#define ON_ENTER_NUM   "{ENTER}" 
#define ON_ESCAPE      "{ESCAPE}" 
#define ON_ESC         "{ESC}" 
#define ON_F1          "{F1}" 
#define ON_F2          "{F2}" 
#define ON_F3          "{F3}" 
#define ON_F4          "{F4}" 
#define ON_F5          "{F5}" 
#define ON_F6          "{F6}" 
#define ON_F7          "{F7}" 
#define ON_F8          "{F8}" 
#define ON_F9          "{F9}" 
#define ON_F10         "{F10}" 
#define ON_F11         "{F11}" 
#define ON_F12         "{F12}" 
#define ON_F13         "{F13}" 
#define ON_F14         "{F14}" 
#define ON_F15         "{F15}" 
#define ON_HELP        "{HELP}" 
#define ON_HOME        "{HOME}" 
#define ON_INSERT      "{INSERT}" 
#define ON_LEFT_ARROW  "{LEFT}" 
#define ON_NUM_LOCK    "{NUMLOCK}" 
#define ON_PAGE_DOWN   "{PGDN}" 
#define ON_PAGE_UP     "{PGUP}" 
#define ON_RETURN      "{RETURN}" 
#define ON_RIGHT_ARROW "{RIGHT}" 
#define ON_SCROLL_LOCK "{SCROLLLOCK}" 
#define ON_TAB         "{TAB}" 
#define ON_UP_ARROW    "{UP}" 

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
	
	// Calls on xlcOnXXX but overwrites existing macros.
	// ???Keep a list of all On macros???
	template<class Key>
	class On {
		using cstr = const char*;
	public:
		On(cstr text, cstr macro)
		{
			Auto<Open> xao([text, macro]() {
				return XExcel<XLOPERX>(Key::On, OPER(text), OPER(macro)) == true;
			});
			Auto<Close> xac([text]() {
				return XExcel<XLOPERX>(Key::On, OPER(text)) == true;
			});
		}
		On(cstr text, cstr macro, bool activate)
		{
			static_assert(std::is_same_v<Key, Sheet>);

			Auto<Open> xao([text, macro, activate]() {
				return XExcel<XLOPERX>(Key::On, OPER(text), OPER(macro), OPER(activate)) == true;
			});
			Auto<Close> xac([text]() {
				return XExcel<XLOPERX>(Key::On, OPER(text)) == true;
			});
		}
		On(const OPER& time, cstr macro, const OPER& tolerance, bool insert)
		{
			static_assert(std::is_same_v<Key, Time>);

			Auto<Open> xao([time, macro, tolerance, insert]() {
				return XExcel<XLOPERX>(Key::On, OPER(time), OPER(macro), OPER(tolerance), OPER(insert)) == true;
			});
			Auto<Close> xac([]() {
				return XExcel<XLOPERX>(Key::On) == true;
			});
		}
	};

} // namespace xll