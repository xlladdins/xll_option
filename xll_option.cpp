// xll_option.cpp - Option value and greeks parameterized by model.
#include <cmath>
//Uncomment to build for versions of Excel prior to 2007
//#define XLOPERX XLOPER
#include "xll_option.h"

using namespace fms;
using namespace xll;

using model = option<variate_base<>>;

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
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option(m, &model::value, f, s, k);
}

AddIn xai_option_delta(
	Function(XLL_DOUBLE, "xll_option_delta", "OPTION.DELTA")
	.Args({
		Arg(XLL_HANDLE, "m", "is a handle to a model."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
	})
	.FunctionHelp("Return the option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option(m, &model::delta, f, s, k);
}

AddIn xai_option_gamma(
	Function(XLL_DOUBLE, "xll_option_gamma", "OPTION.GAMMA")
	.Args({
		Arg(XLL_HANDLE, "m", "is a handle to a model."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
	})
	.FunctionHelp("Return the option gamma.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_gamma(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option(m, &model::gamma, f, s, k);
}

AddIn xai_option_vega(
	Function(XLL_DOUBLE, "xll_option_vega", "OPTION.VEGA")
	.Args({
		Arg(XLL_HANDLE, "m", "is a handle to a model."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
	})
	.FunctionHelp("Return the option vega.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_vega(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option(m, &model::vega, f, s, k);
}

AddIn xai_option_implied(
	Function(XLL_DOUBLE, "xll_option_implied", "OPTION.IMPLIED")
	.Args({
		Arg(XLL_HANDLE, "m", "is a handle to a model."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "v", "is the option value."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
		Arg(XLL_DOUBLE, "s0", "is the initial vol guess. Default is 0.1"),		
		Arg(XLL_WORD,   "n", "is the maximum number of iterations. Default is 10."),
		Arg(XLL_DOUBLE, "eps", "is value precision. Default is sqrt of machine epsilon."),
	})
	.FunctionHelp("Return the option implied.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_implied(HANDLEX m, double f, double v, double k, double s0, WORD n, double eps)
{
#pragma XLLEXPORT
	return xll_option(m, &model::implied, f, v, k, s0, static_cast<size_t>(n), eps);
}
