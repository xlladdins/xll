// on.h - xlcOnXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "auto.h"
#include "excel.h"

// For use with On<Key>
#define ON_SHIFT       X_("+")
#define ON_CTRL        X_("^")
#define ON_ALT         X_("%")
#define ON_COMMAND     X_("*")
#define ON_ENTER       X_("~")
#define ON_BACKSPACE   X_("{BACKSPACE}") 
#define ON_BS          X_("{BS}") 
#define ON_BREAK       X_("{BREAK}") 
#define ON_CAPS_LOCK   X_("{CAPSLOCK}") 
#define ON_CLEAR       X_("{CLEAR}") 
#define ON_DELETE      X_("{DELETE}") 
#define ON_DEL         X_("{DEL}") 
#define ON_DOWN_ARROW  X_("{DOWN}") 
#define ON_END         X_("{END}") 
#define ON_ENTER_NUM   X_("{ENTER}") 
#define ON_ESCAPE      X_("{ESCAPE}") 
#define ON_ESC         X_("{ESC}") 
#define ON_F1          X_("{F1}") 
#define ON_F2          X_("{F2}") 
#define ON_F3          X_("{F3}") 
#define ON_F4          X_("{F4}") 
#define ON_F5          X_("{F5}") 
#define ON_F6          X_("{F6}") 
#define ON_F7          X_("{F7}") 
#define ON_F8          X_("{F8}") 
#define ON_F9          X_("{F9}") 
#define ON_F10         X_("{F10}") 
#define ON_F11         X_("{F11}") 
#define ON_F12         X_("{F12}") 
#define ON_F13         X_("{F13}") 
#define ON_F14         X_("{F14}") 
#define ON_F15         X_("{F15}") 
#define ON_HELP        X_("{HELP}") 
#define ON_HOME        X_("{HOME}") 
#define ON_INSERT      X_("{INSERT}") 
#define ON_LEFT_ARROW  X_("{LEFT}") 
#define ON_NUM_LOCK    X_("{NUMLOCK}") 
#define ON_PAGE_DOWN   X_("{PGDN}") 
#define ON_PAGE_UP     X_("{PGUP}") 
#define ON_RETURN      X_("{RETURN}") 
#define ON_RIGHT_ARROW X_("{RIGHT}") 
#define ON_SCROLL_LOCK X_("{SCROLLLOCK}") 
#define ON_TAB         X_("{TAB}") 
#define ON_UP_ARROW    X_("{UP}") 

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
		using cstr = const char*;
	public:
		On(cstr text, cstr macro)
		{
			Auto<Open> xao([text, macro]() {
				return ExcelX(Key::On, OPER(text), OPER(macro)) == true;
			});
			Auto<Close> xac([text]() {
				return ExcelX(Key::On, OPER(text)) == true;
			});
		}
		On(cstr text, cstr macro, bool activate)
		{
			static_assert(std::is_same_v<Key, Sheet>);

			Auto<Open> xao([text, macro, activate]() {
				return ExcelX(Key::On, OPER(text), OPER(macro), OPER(activate)) == true;
			});
			Auto<Close> xac([text]() {
				return ExcelX(Key::On, OPER(text)) == true;
			});
		}
		On(const OPER& time, cstr macro, const OPER& tolerance, bool insert)
		{
			static_assert(std::is_same_v<Key, Time>);

			Auto<Open> xao([time, macro, tolerance, insert]() {
				return ExcelX(Key::On, OPER(time), OPER(macro), OPER(tolerance), OPER(insert)) == true;
			});
			Auto<Close> xac([]() {
				return ExcelX(Key::On) == true;
			});
		}
	};
} // namespace xll