// document.cpp - generate add-in documentaton
#include <fstream>
#include "../splitpath.h"
#include "xll_macrofun.h"

namespace xll {

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
		border-padding: 5px;
		border-collapse: collapse;
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
			std::string functionText
				= arg.FunctionText().to_string();

			// factor out!!!
			splitpath sp(Excel4(xlGetName).to_string().c_str());
			std::string ofile(sp.dirname() + functionText + ".html");
			std::ofstream ofs(ofile, std::ios::out);

ofs << R"xyzyx(<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
)xyzyx"
    << html_style_css
    << "<title>" << functionText << "</title>\n"
    << html_head_post;

			std::string macroType(" ");
			macroType.append(arg.Type().to_string());
ofs << "<body>\n\t<h1>" << functionText << macroType << "</h1>\n\t"
    << "<p>This article describes the formula syntax of the " 
    << functionText << macroType << "</p>\n\t";

			std::string functionHelp
				= arg.FunctionHelp().to_string();
ofs << "<h2>Description</h2>\n\t<p>\n" << functionHelp << "\n\t</p>\n\t";

			std::string argumentText
				= arg.ArgumentText().to_string();
ofs << "<h2>Syntax</h2>\n\t<p>" << functionText << "(" << argumentText << ")</p>\n\t"
    << "<blockquote>\n\t<table>\n\t<tbody>\n\t";

			for (unsigned i = 0; i < arg.ArgumentCount(); ++i) {
ofs << "\t<tr>\n\t\t"
	<< "<td>" << arg.ArgumentName(i).to_string() << "</td>\n\t\t"
	<< "<td>" << arg.ArgumentHelp(i).to_string() << "</td>\n\t\t"
	<< "</tr>\n\t";
			}
ofs << "</tbody>\n</table>\n</blockquote>\n"
	<< "<p>" << arg.Documentation() << "</p>\n"
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
	bool Documentation(const char* category, const char* description)
	{
		splitpath sp(Excel4(xlGetName).to_string().c_str());
		std::string ofile(sp.dirname() + "index.html");
		std::ofstream ofs(ofile, std::ios::out);

		// sort by Category then by FunctionText
		std::map<OPER, std::map<OPER, OPER>> cat_text; // Category -> FunctionText -> FunctionHelp

		try {
			for (auto& [key, arg] : AddIn::Map) {
				if (arg.Documentation().length() != 0) {
					cat_text[arg.Category()][arg.FunctionText()] = arg.FunctionHelp();
					Document(arg);
				}
			}

			// index.html preamble
ofs << R"(<!DOCTYPE html>
<head>
    <meta charset="UTF-8" />
	)"
	<< html_style_css
	<< "<title>" << category << "</title>"
	<< R"(
</head>
<body>
<h1>)"
<< category
<< R"(</h1>
	<p>
		Functions and macros of the )" << category << R"( add-in.
	</p>
	<p>
	)"
<< description << "</p>\n";

			// categories and all functions belonging the to each category
			for (const auto& [cat, text_help] : cat_text) {
ofs << "<h2>Category " << cat.to_string() << "</h2>\n"
	<< "<table>\n<tbody>\n";
				for (const auto& [text, help] : text_help) {
					std::string functionText = text.to_string();
					std::string functionHelp = help.to_string();
					std::string href = functionText;
					//!!! use isHandle ???
					if (href[0] == '\\') {
						href.erase(0, 1);
					}
ofs << "<tr>\n\t"
	<< "<td><a href=\"" << href << ".html\">" << functionText << "</a></td>\n"
	<< "<td>" << functionHelp << "</td>\n"
	<< "</tr>\n";
				}
ofs << "</tbody>\n</table>\n";
			}
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());

			return false;
		}

		return true;
	}
}
