// register.h - register an add-in function or macro
#pragma once
#include "xll.h"

namespace xll {
	struct Register {
        LPXLOPER12 ModuleText; 
        LPXLOPER12 Procedure;
        LPXLOPER12 TypeText; 
        LPXLOPER12 FunctionText;
        LPXLOPER12 ArgumentText; 
        LPXLOPER12 MacroType;
        LPXLOPER12 Category; 
        LPXLOPER12 ShortcutText;
        LPXLOPER12 HelpTopic; 
        LPXLOPER12 FunctionHelp;
        LPXLOPER12 ArgumentHelp1; 
        // ...
	};
}