// xll.h - Excel add-in interface
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "defines.h"
#include "error.h"
#include "exports.h"
#include "addin.h"
#include "register.h"
#include "on.h"
#include "handle.h"
#include "fp.h"
#include "platform.h"

namespace xll {
	bool Documentation(const char* category, const char* description);
}