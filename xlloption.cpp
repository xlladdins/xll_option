// xlloption.cpp - Sample xll project.
#include <cmath>
//Uncomment to build for versions of Excel prior to 2007
//#define XLOPERX XLOPER
#include "xlloption.h"
#include "../../keithalewis/fmsoption/fms_normal.h"

//using obase = fms::option_base<double, double, double>;

using namespace fms;
using namespace xll;

using option_model_normal = fms::option<fms::variate::normal<>>;

AddIn xai_option_model_normal(
	Function(XLL_HANDLE, "xll_option_model_normal", "OPTION.MODEL.NORMAL")
	.Category("Option")
	.FunctionHelp("Return handle to normal option model.")
	.Uncalced()
);
HANDLEX WINAPI xll_option_model_normal()
{
#pragma XLLEXPORT
	handle<variate_base<>> m(new variate::normal<>{});

	return m.get();
}

AddIn xai_option_value(
	Function(XLL_DOUBLE, "xll_option_value", "OPTION.VALUE")
	.Args({
		Arg(XLL_HANDLE, "m", "is a handle to a model."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
	})
	.FunctionHelp("Return the option value.")
	.Category("Option")
	//.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
);
double WINAPI xll_option_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_value<variate::normal<>>(m, f, s, k);
}
