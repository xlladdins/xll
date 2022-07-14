// error.h - Error reporting functions
// Copyright (c) KALX, LLC. All rights reserved.
#pragma once
//#include <Windows.h>

enum /*!!! class*/ { 
	XLL_ALERT_ERROR   = 1,
	XLL_ALERT_WARNING = 2, 
	XLL_ALERT_INFO    = 4,
//	XLL_ALERT_LOG     = 8	// turn on logging (should be separate)
};

/// Set error level and return old
unsigned long XLL_ALERT_LEVEL(unsigned long level);

/// OKCANCEL message box. Cancel turns off error bit
int XLL_ERROR(const char* e, bool force = false);

/// OKCANCEL message box. Cancel turns off warning bit
int XLL_WARNING(const char* e, bool force = false);

/// OKCANCEL message box. Cancel turns off information bit
int XLL_INFO(const char* e, bool force = false);
