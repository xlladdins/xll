// document.h - Generate HTML documentation for an add-in
#pragma once

namespace xll {

	bool HTMLDocumentation(const char*, const char*);

	// call to generate all documentation for an add-in
	inline int Documentation([[maybe_unused]] const char* name, [[maybe_unused]] const char* description)
	{
#ifdef _DEBUG
		Auto<OpenAfter> aoa_document([name, description]() { return HTMLDocumentation(name, description); });
#endif		
		return 1;
	}


}
