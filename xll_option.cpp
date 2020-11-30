// xll_option.cpp - Option value and greeks parameterized by model.
//Uncomment to build for versions of Excel prior to 2007
//#define XLOPERX XLOPER
#include "fmsoption/fms_option.h"
#include "fmsoption/fms_variate_handle.h"
#include "xll/xll/xll.h"

using namespace fms;
using namespace xll;

AddIn xai_option_value(
	Function(XLL_DOUBLE, "xll_option_value", "OPTION.VALUE")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
	})
	.FunctionHelp("Return the call (k > 0) or put (k < 0) option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.value(f, s, k);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}
AddIn xai_option_call_value(
	Function(XLL_DOUBLE, "xll_option_call_value", "OPTION.CALL.VALUE")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the call strike."),
	})
	.FunctionHelp("Return the call option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_call_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.value(f, s, payoff::call(k));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}
AddIn xai_option_put_value(
	Function(XLL_DOUBLE, "xll_option_put_value", "OPTION.PUT.VALUE")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
	})
	.FunctionHelp("Return the put option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_put_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.value(f, s, payoff::put(k));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}

AddIn xai_option_delta(
	Function(XLL_DOUBLE, "xll_option_delta", "OPTION.DELTA")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
		.FunctionHelp("Return the call (k > 0) or put (k < 0) option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.delta(f, s, k);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}
AddIn xai_option_call_delta(
	Function(XLL_DOUBLE, "xll_option_call_delta", "OPTION.CALL.DELTA")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
		.FunctionHelp("Return the call option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_call_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.delta(f, s, payoff::call(k));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}
AddIn xai_option_put_delta(
	Function(XLL_DOUBLE, "xll_option_put_delta", "OPTION.PUT.DELTA")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
		.FunctionHelp("Return the put option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_put_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.delta(f, s, payoff::put(k));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}

AddIn xai_option_gamma(
	Function(XLL_DOUBLE, "xll_option_gamma", "OPTION.GAMMA")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
		.FunctionHelp("Return the option gamma.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_gamma(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.gamma(f, s, k);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}

AddIn xai_option_vega(
	Function(XLL_DOUBLE, "xll_option_vega", "OPTION.VEGA")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the volatility."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
		.FunctionHelp("Return the option vega.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_vega(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.vega(f, s, k);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}

AddIn xai_option_implied(
	Function(XLL_DOUBLE, "xll_option_implied", "OPTION.IMPLIED")
	.Args({
		Arg(XLL_HANDLE, "m", "is a handle to a model."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "v", "is the option value."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
		Arg(XLL_DOUBLE, "s", "is the initial vol guess. Default is 0.1"),		
		Arg(XLL_WORD,   "n", "is the maximum number of iterations. Default is 10."),
		Arg(XLL_DOUBLE, "eps", "is value precision. Default is sqrt of machine epsilon."),
	})
	.FunctionHelp("Return the option implied vol.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_implied(HANDLEX m, double f, double v, double k, double s, WORD n, double eps)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.implied(f, v, std::abs(k), s, n, eps);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return std::numeric_limits<double>::quiet_NaN();
}
