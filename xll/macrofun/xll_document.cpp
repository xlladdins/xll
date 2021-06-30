// document.cpp - generate add-in documentaton
#include <fstream>
#include <filesystem>
#include "xll_macrofun.h"
#include "../splitpath.h"

namespace xll {

	// Args data as xml
	inline std::ostream& to_xml(std::ostream& os, const Args& args,
		const std::string& xslt = "xll.xslt")
	{
		os << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
		os << "<?xml-stylesheet href=\"" << xslt << "\" type=\"text/xsl\"?>\n";

		// <key>val</key>
		auto element = [&os](const OPER& key, const OPER& val) {
			os << "<" << key.to_string() << ">";
			if (val.is_str()) {
				os << val.to_string();
			}
			os << "</" << key.to_string() << ">\n";
		};

		os << "<args>\n";

		for (const auto& [key, val] : args.argMap) {
			element(key, val);
		}

		// arguments
		const OPER& type = args["argumentType"];
		const OPER& name = args["argumentName"];
		const OPER& help = args["argumentHelp"];
		const OPER& init = args["argumentDefault"];
		
		os << "<arguments>\n";
		for (unsigned i = 0; i < type.size(); ++i) {
			os << "<arg>\n";
			element(OPER("type"), type[i]);
			element(OPER("name"), name[i]);
			element(OPER("help"), help[i]);
			element(OPER("init"), init[i]);
			os << "</arg>\n";
		}
		os << "</arguments>\n";

		element(OPER("functionTextSafe"), args.FunctionText().safe());
		element(OPER("type"), args.Type());
		os << "<documentation>\n";
		os << "<![CDATA[" << args.documentation << "]]>";
		os << "</documentation>\n";

		os << "</args>\n";

		return os;
	}

	// AddIn data as xml
	inline std::ostream& to_xml(std::ostream& os, const AddIn& ai)
	{
		for (const auto& [key, arg] : ai.Map) {
			// os << ...
		}

		return os;
	}

	// Write html documentation given Args.
	bool Document(const Args& args)
	{
		try {
			OPER functionText = args.FunctionText();
			std::string safe = functionText.safe().to_string();

			// factor out!!!
			path sp(Excel4(xlGetName).to_string().c_str());
			std::string ofile(sp.dirname() + safe + ".xml");
			std::ofstream ofs(ofile, std::ios::out);
			to_xml(ofs, args);
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());

			return false;
		}

		return true;
	}

	// Generate documentation for add-ins;
	bool Documentation(const char* module, const char* description)
	{
		path<char> sp;
		sp.split(Excel4(xlGetName).to_string().c_str());

		for (auto& [key, arg] : AddIn::Map) {
			if (arg.Documentation().length() != 0) {
				std::string file = key.safe().to_string();
				std::ofstream xml(sp.dirname() + file + ".xml", std::ios::out);
				to_xml(xml, arg);
			}
		}

		std::string ofile(sp.dirname() + "index.html");
		std::ofstream ofs(ofile, std::ios::out);

		// sort by Category, Type, FunctionText
		std::map<OPER, std::map<OPER, std::map<OPER, Args*>>> cat_type_text;
		
		try {
			for (auto& [key, arg] : AddIn::Map) {
				if (arg.Documentation().length() != 0) {
					OPER cat = arg.Category();
					OPER type = arg.Type();
					OPER text = arg.FunctionText();
					if (arg.isHandle()) {
						// strip off leading '\\'
						text = Excel(xlfRight, text, OPER(text.val.str[0] - 1));
					}
					cat_type_text[cat][type][text] = &arg;
					Document(arg);
				}
			}

			// index.html preamble
			ofs << R"(<!DOCTYPE html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <link href="xll.css" rel="stylesheet" type="text/css" />
)";
	ofs << "<title>" << module << "</title>"
	<< R"(
</head>
<body>
<h1>)"
<< module
<< R"(</h1>
	<p>
		Functions and macros of the )" << module << R"( add-in.
	</p>
	<p>
	)"
<< description << "</p>\n";

			// categories and all functions belonging the to each category
			for (const auto& [cat, type_text_parg] : cat_type_text) {
				for (const auto& [type, text_parg] : type_text_parg) {
					ofs << "<h2>" << cat.to_string() << " " << type.to_string() << "s</h2>\n"
						<< "<table>\n<tbody>\n";
					for (const auto& [text, parg] : text_parg) {
						std::string functionText = parg->FunctionText().to_string();
						std::string functionHelp = parg->FunctionHelp().to_string();
						std::string href = parg->FunctionText().safe().to_string();
						ofs << "<tr>\n\t"
							<< "<td><a href=\"" << href << ".xml\">" << functionText << "</a></td>\n"
							<< "<td>" << functionHelp << "</td>\n"
							<< "</tr>\n";
					}
					ofs << "</tbody>\n</table>\n";
				}
			}
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());

			return false;
		}

		return true;
	}
}
