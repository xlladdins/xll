// addin.h - convenience wrapper for Excel add-ins
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <cwctype>
#include <map>
#include "auto.h"
#include "args.h"

namespace xll {

	/// Manage the lifecycle of an Excel add-in.
	class AddIn {
    public:
        static inline std::map<OPER, Args> KeyArgsMap;
        static inline std::map<double, OPER> RegIdKeyMap;

        /// Register and Unregister an add-in when Excel calls xlAutoOpen and xlAutoClose.
        AddIn(const Args& args)
        {
            for (const OPER& key : args.Key()) {
                Auto<Open> ao([=]() {
                    try {
                        OPER oReg = args.Register();
                        RegIdKeyMap.insert(std::pair{ oReg.val.num, key });
                        KeyArgsMap.insert(std::pair{ key, args });

                        return TRUE;
                    }
                    catch (const std::exception& ex) {
                        MessageBoxA(GetForegroundWindow(), ex.what(), "AddIn Auto<Open> failed", MB_OK | MB_ICONERROR);

                        return FALSE;
                    }
                });
                
                Auto<Close> ac([=]() {
                    try {
                        args.Unregister();
                    }
                    catch (const std::exception& ex) {
                        MessageBoxA(GetForegroundWindow(), ex.what(), "AddIn Auto<Close> failed", MB_OK | MB_ICONERROR);

                        return FALSE;
                    }

                    return TRUE;
                });
            }
        }
    };
	using AddIn12 = AddIn;
	using AddInX = AddIn;
} // xll namespace


