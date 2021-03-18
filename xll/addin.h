// addin.h - convenience wrapper for Excel add-ins
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <map>
#include "args.h"

namespace xll {

    /// <summary>
    /// Store argument add-in data for xlfRegister
    /// </summary>
	struct AddIn {
        // add-ins indexed by Excel name
        static inline std::map<OPER, Args> Map;
 
        AddIn(const Args& args) noexcept
        {
            const auto& key = args.FunctionText();
            if (Map.find(key) != Map.end()) {
                std::basic_string<TCHAR> msg{ TEXT("AddIn previously defined: ") };
                msg.append(key.val.str + 1, key.val.str[0]);
                MessageBox(GetForegroundWindow(), msg.c_str(), 0, MB_OK);
            }
            Map[key] = args;
        }

        static size_t Erase(const OPER& key)
        {
            return Map.erase(key);
        }

        // Get arguments using Excel function text name.
        static Args* Arguments(const OPER& name)
        {
            auto i = Map.find(name);
            
            return i != Map.end() ? &i->second : nullptr;
        }
        // Get arguments using register id.
        static Args* Arguments(double regid)
        {
            for (auto& args : Map) {
                if (args.second.RegisterId().val.num == regid) {
                    return &args.second;
                }
            }

            return nullptr;
        }
    };

} // xll namespace


