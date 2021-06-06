// document.cpp - generate add-in documentaton
#include <fstream>
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

		for (const auto& [key, val] : args.argMap) {
			element(key, val);
		}
		// arguments
		const OPER& type = args["argumentType"];
		const OPER& name = args["argumentName"];
		const OPER& desc = args["argumentDesc"];
		const OPER& init = args["argumentDefault"];
		os << "<arguments>\n";
		for (unsigned i = 0; i < type.size(); ++i) {
			os << "<arg>\n";
			element(OPER("type"), type);
			element(OPER("name"), name);
			element(OPER("desc"), desc);
			element(OPER("init"), init);
			os << "</arg>\n";
		}
		os << "</arguments>\n";

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

	// CSS style
	static const char* html_style_css = R"xyzyx(
	<style>
	body{
		background-color: #fff;
		color: #363636;
		font-family: 'Segoe UI',Calibri,Arial,'Helvetica Neue',Verdana,Helvetica,Sans-Serif;
		margin: 0.5in;
		padding: 0;
	}
	table {
		text-align: left;
		padding: 5px;
		border-collapse: collapse;
	}
	th {
		color: white;
		background-color: #00a1f1;
	}
	th:first-child, td:first-child {
		padding-right: 1em;
		text-align: right;
		font-weight: bold;
	}
	tbody>tr:nth-child(odd) {
		background-color: #f2f2f2;
	}
	td:first-child {
		background-color: #fff;
	}
	.katex{
		font-size: 1em !important;
		font-weight: 200 !important;
	}
	</style>
)xyzyx";

	// load katex
	inline const char* html_head_post = R"xyzyx(
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/katex.min.css" 
		integrity="sha384-AfEj0r4/OFrOo5t7NnNe46zW/tFgW6x/bCJG8FqQCEo3+Aro6EYUG4+cU+KJWu/X" crossorigin="anonymous">
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/katex.min.js" 
		integrity="sha384-g7c+Jr9ZivxKLnZTDUhnkOnsh30B4H0rpLUpJ4jAIKs4fnJI+sEnkvrMWph2EDg4" crossorigin="anonymous"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/contrib/auto-render.min.js" 
		integrity="sha384-mll67QQFJfxn0IYznZYonOWZ644AWYC+Pt2cHqMaRhXVrursRwvLnLaebdGIlYNa" crossorigin="anonymous"
        onload="renderMathInElement(document.body, {fleqn: true});"></script>
</head>
)xyzyx";

	// Write html documentation given Args.
	bool Document(const Args& arg)
	{
		try {
			OPER functionText = arg.FunctionText();
			std::string safe = functionText.safe().to_string();

			// factor out!!!
			path sp(Excel4(xlGetName).to_string().c_str());
			std::string ofile(sp.dirname() + safe + ".html");
			std::ofstream ofs(ofile, std::ios::out);

ofs << R"xyzyx(<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
)xyzyx"
    << html_style_css
    << "<title>" << functionText.to_string() << "</title>\n"
    << html_head_post;

			std::string macroType(" ");
			macroType.append(arg.Type().to_string());
ofs << "<body>\n\t<h1>" << functionText.to_string() << macroType << "</h1>\n\t"
    << "<p>This article describes the formula syntax of the " 
    << functionText.to_string() << macroType << "</p>\n\t";

			std::string functionHelp
				= arg.FunctionHelp().to_string();
ofs << "<h2>Description</h2>\n\t<p>\n" << functionHelp << "\n\t</p>\n\t";

			std::string argumentText
				= arg.ArgumentText().to_string();
			if (arg.isFunction()) {
				ofs << "<h2>Syntax</h2>\n\t<p>" << functionText.to_string() << "(" << argumentText << ")</p>\n\t"
					<< "<blockquote>\n\t<table>\n\t<tbody>\n\t";

				for (unsigned i = 1; i <= arg.ArgumentCount(); ++i) {
					ofs << "\t<tr>\n\t\t"
						<< "<td>" << arg.ArgumentName(i).to_string() << "</td>\n\t\t"
						<< "<td>" << arg.ArgumentHelp(i).to_string() << "</td>\n\t\t"
						<< "</tr>\n\t";
				}
				ofs << "</tbody>\n</table>\n</blockquote>\n";
			}
	ofs << "<p>" << arg.Documentation() << "</p>\n"
	<< R"(
	<footer>
		Return to <a href="index.html">index</a>.
	</footer>
</body>
</html>
)";
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
    <meta charset="UTF-8" />
	)"
	<< html_style_css
	<< "<title>" << module << "</title>"
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
							<< "<td><a href=\"" << href << ".html\">" << functionText << "</a></td>\n"
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
