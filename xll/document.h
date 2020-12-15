// document.h - Generate HTML documentation for an add-in
#pragma once
#include "addin.h"
#include "error.h"
#include "xllio.h"

namespace xll {

	inline const OPER top("<!DOCTYPE HTML>\n<HTML>\n");
	inline const OPER head("<HEAD>\n\t<TITLE>FunctionText</TITLE>\n</HEAD>\n<BODY>");
	inline const OPER bot("</BODY>\n</HTML>");

	// Write html documentation given Excel function text.
	inline void Document(const OPER& ft)
	{
		auto sub = [](const OPER& t, const OPER& o, const OPER& n) { return Excel(xlfSubstitute, t, o, n);  };

		try {
			OPER dir = dirname(Excel(xlGetName));
			xlfFile html(dir & ft & OPER(".html"), 3);

			html.write(top);
			html.write(sub(head, OPER("FunctionText"), ft));

			const Args& arg = AddIn::Args(ft);
			const OPER& mt = arg.MacroType();
			OPER type = OPER(mt == 1 ? " function" : mt == 2 ? " macro" : "");
			html.write(OPER("<H1>") & ft & type & OPER("</H1>\n"));

			html.write(OPER("<P>\nThis article describes the formula syntax and usage of the "));
			html.write(ft);
			html.write(type);
			html.write(OPER(" in Microsoft Excel.\n</P>\n"));

			html.write(OPER("<H2>Description</H2>\n<P>\n"));
			html.write(arg.FunctionHelp());
			html.write(OPER("\n</P>\n"));

			html.write(OPER("<H2>Syntax</H2>\n<P>\n"));
			html.write(ft);
			html.write(OPER("("));
			html.write(arg.ArgumentText());
			html.write(OPER(")\n</P>\n"));

			if (mt == 1) {
				const auto& help = arg.ArgumentHelp();
				if (help.size() > 0) {
					const auto& name = arg.ArgumentName();
					html.write(OPER("<P>\nThe "));
					html.write(ft);
					html.write(OPER(" function syntax has the following arguments:</P>\n<UL>\n"));
					for (int i = 0; i < help.size(); ++i) {
						html.write(OPER("<LI><B>"));
						html.write(name[i]);
						html.write(OPER("</B> "));
						html.write(help[i]);
						html.write(OPER("</LI>\n"));
					}
					html.write(OPER("</UL>\n</P>\n"));
				}

			}
			const auto& doc = arg.Documentation();
			if (doc.length() > 0) {
				html.write(OPER("<H2>Remarks</H2>\n<P>\n"));
				html.write(OPER(doc.c_str()));
				html.write(OPER("\n</P>\n"));
			}

			html.write(bot);
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());
		}
	}

}
