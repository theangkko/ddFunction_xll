// xll_template.cpp - Sample xll project.
#include <cmath> // for double tgamma(double)
#include "main.h"
#include "steamprop.h"
#include "IF97.h"


using namespace xll;
using namespace IF97;


#ifdef _DEBUG

// Create XML documentation and index.html in `$(TargetPath)` folder.
// Use `xsltproc file.xml -o file.html` to create HTML documentation.
//Auto<Open> xao_template_docs([]() {
//
//	return Documentation("MATH", "Documentation for the xll_template add-in.");
//
//});

#endif // _DEBUG




AddIn xai_tgamma(
	// Return double, C++ name of function, Excel name.
	Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
	// Array of function arguments.
	.Arguments({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
	})
	// Function Wizard help.
	.FunctionHelp("Return the Gamma function value.")
	// Function Wizard category.
	.Category("MATH")
	// URL linked to `Help on this function`.
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
	.Documentation(R"xyzyx(
The <i>Gamma</i> function is \(\Gamma(x) = \int_0^\infty t^{x - 1} e^{-t}\,dt\), \(x \ge 0\).
If \(n\) is a natural number then \(\Gamma(n + 1) = n! = n(n - 1)\cdots 1\).
<p>
Any valid HTML using <a href="https://katex.org/" target="_blank">KaTeX</a> can 
be used for documentation.
)xyzyx")
);
// WINAPI calling convention must be specified
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT // must be specified to export function

	return tgamma(x);
}



// Press Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://xlladdins.github.io/Excel4Macros/
AddIn xai_macro(
	// C++ function, Excel name of macro
	Macro("xll_macro", "XLL.MACRO")
);
// Macros must have `int WINAPI (*)(void)` signature.
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	// https://xlladdins.github.io/Excel4Macros/reftext.html
	// A1 style instead of default R1C1.
	OPER reftext = Excel(xlfReftext, Excel(xlfActiveCell), OPER(true));
	// UTF-8 strings can be used.
	Excel(xlcAlert, OPER("XLL.MACRO called with активный 细胞: ") & reftext);

	return TRUE;
}




////////


AddIn xai_stmdensity(
	// Return double, C++ name of function, Excel name.
	Function(XLL_DOUBLE, "ddstmdensity", "dd.STM")
	// Array of function arguments.
	.Arguments({
		Arg(XLL_DOUBLE, "p", "is the value for which you want to calculate Gamma."),
		Arg(XLL_DOUBLE, "t", "is the value for which you want to calculate Gamma.")
		})
	// Function Wizard help.
	.FunctionHelp("Return the TEST function value.")
	// Function Wizard category.
	.Category("MATH")
	// URL linked to `Help on this function`.
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
	.Documentation(R"xyzyx(
The <i>Gamma</i> function is \(\Gamma(x) = \int_0^\infty t^{x - 1} e^{-t}\,dt\), \(x \ge 0\).
If \(n\) is a natural number then \(\Gamma(n + 1) = n! = n(n - 1)\cdots 1\).
<p>
Any valid HTML using <a href="https://katex.org/" target="_blank">KaTeX</a> can 
be used for documentation.
)xyzyx")
);
// WINAPI calling convention must be specified
double WINAPI ddstmdensity(double p, double t)
{
#pragma XLLEXPORT // must be specified to export function

	return stmdensity(p, t);
}



AddIn xai_stmPTH(
	// Return double, C++ name of function, Excel name.
	Function(XLL_DOUBLE, "stmPTH", "dd.STMPTH")
	// Array of function arguments.
	.Arguments({
		Arg(XLL_DOUBLE, "p", "is the value for which you want to calculate Gamma."),
		Arg(XLL_DOUBLE, "t", "is the value for which you want to calculate Gamma.")
		})
	// Function Wizard help.
	.FunctionHelp("Return the TEST function value.")
	// Function Wizard category.
	.Category("MATH")
	// URL linked to `Help on this function`.
	.HelpTopic("https://theangkko.github.io")
	.Documentation(R"xyzyx(Sample Explanation)xyzyx")
);
// WINAPI calling convention must be specified
double WINAPI stmPTH(double p, double t)
{
#pragma XLLEXPORT // must be specified to export function

	return hmass_Tp(t, p)/1000;
}




