// xll_option.cpp - Option value and greeks parameterized by model.
//Uncomment to build for versions of Excel prior to 2007
#include "xll_option.h"

using namespace fms;
using namespace fms::variate;
using namespace xll;

#ifdef _DEBUG
Auto<OpenAfter> xaoa_option_doc([]() {
	return xll::Documentation("OPTION", R"(
European option pricing.
)");
});
#endif // _DEBUG


AddIn xai_option_moneyness(
	Function(XLL_DOUBLE, "xll_option_moneyness", "OPTION.MONEYNESS")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a model.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the strike.", "100"),
		})
		.FunctionHelp("Return the option moneyness.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzyx(
Moneyness \(m = m(f,s,k)\) is defined by \(F = k\) if and only if \(X = m\) where
\(F = fe^{s X - \kappa(s)}\) is the underlying price at expiration and
\(\kappa(s)\) is the cumulant \(\kappa(s) = \log E[e^{sX}]\). 
)xyzyx")
);
double WINAPI xll_option_moneyness(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.moneyness(f, s, k);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return XLL_NAN;
}


AddIn xai_option_value(
	Function(XLL_DOUBLE, "xll_option_value", "OPTION.VALUE")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the strike.", "100"),
	})
	.FunctionHelp("Return the call (k > 0) or put (k < 0) option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return the call value \(E[\max\{F - k, 0\}]\) if \(k\) is positive 
or the put value \(E[\max\{|k| - F, 0\}]\) if \(k\) is negative.
)xyzx")
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

	return XLL_NAN;
}
AddIn xai_option_call_value(
	Function(XLL_DOUBLE, "xll_option_call_value", "OPTION.CALL.VALUE")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the call strike.", "100"),
	})
	.FunctionHelp("Return the call option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return the call value \(E[\max\{F - k, 0\}]\) if \(k\). 
)xyzx")
);
double WINAPI xll_option_call_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_value(m, f, s, k);
}
AddIn xai_option_put_value(
	Function(XLL_DOUBLE, "xll_option_put_value", "OPTION.PUT.VALUE")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the strike.", "100"),
	})
	.FunctionHelp("Return the put option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return put value \(E[\max\{k - F, 0\}]\).
)xyzx")
);
double WINAPI xll_option_put_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_value(m, f, s, -k);
}

AddIn xai_option_delta(
	Function(XLL_DOUBLE, "xll_option_delta", "OPTION.DELTA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the put strike.", "100"),
		})
		.FunctionHelp("Return the call (k > 0) or put (k < 0) option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return the derivative of option value with respect to the forward \(f\).
)xyzx")
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

	return XLL_NAN;
}
AddIn xai_option_call_delta(
	Function(XLL_DOUBLE, "xll_option_call_delta", "OPTION.CALL.DELTA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the put strike.", "100"),
		})
	.FunctionHelp("Return the call option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return the derivative of call value with respect to the forward \(f\).
)xyzx")
);
double WINAPI xll_option_call_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_delta(m, f, s, k);
}
AddIn xai_option_put_delta(
	Function(XLL_DOUBLE, "xll_option_put_delta", "OPTION.PUT.DELTA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the put strike.", "100"),
		})
	.FunctionHelp("Return the put option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_put_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_delta(m, f, s, -k);
}

AddIn xai_option_gamma(
	Function(XLL_DOUBLE, "xll_option_gamma", "OPTION.GAMMA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the put strike.", "100"),
		})
	.FunctionHelp("Return the option gamma.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return the second derivative of option value with respect to the forward \(f\).
)xyzx")
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

	return XLL_NAN;
}

AddIn xai_option_vega(
	Function(XLL_DOUBLE, "xll_option_vega", "OPTION.VEGA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "s", "is the vol.", "0.1"),
		Arg(XLL_DOUBLE, "k", "is the put strike.", "100"),
		})
	.FunctionHelp("Return the option vega.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return the derivative of option value with respect to the vol \(s\).
)xyzx")
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

	return XLL_NAN;
}

AddIn xai_option_implied(
	Function(XLL_DOUBLE, "xll_option_implied", "OPTION.IMPLIED")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a model.", "\"=\\VARIATE.NORMAL()\""),
		Arg(XLL_DOUBLE, "f", "is the forward.", "100"),
		Arg(XLL_DOUBLE, "v", "is the option value.", "3.98"),
		Arg(XLL_DOUBLE, "k", "is the strike.", "100"),
		Arg(XLL_DOUBLE, "s", "is the initial vol guess. Default is 0.1", "0.1"),		
		Arg(XLL_WORD,   "n", "is the maximum number of iterations. Default is 10."),
		Arg(XLL_DOUBLE, "eps", "is value precision. Default is sqrt of machine epsilon."),
	})
	.FunctionHelp("Return the option implied vol.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
	.Documentation(R"xyzx(
Return the vol \(s\) that recovers the option value.
)xyzx")
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

	return XLL_NAN;
}

AddIn xai_option_improve(
	Function(XLL_DOUBLE, "xll_option_improve", "OPTION.IMPROVE")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a model."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "v", "is the option value."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
		Arg(XLL_DOUBLE, "s", "is the initial vol guess. Default is 0.1"),
		})
		.FunctionHelp("Return the option improved vol.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_improve(HANDLEX m, double f, double v, double k, double s)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			return o.improve(s, f, v, k);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return XLL_NAN;
}

AddIn xai_option_digital_value(
	Function(XLL_DOUBLE, "xll_option_digital_value", "OPTION.DIGITAL.VALUE")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
		})
	.FunctionHelp("Return the digital call (k > 0) or put (k < 0) option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			if (k >= 0) {
				return o.value(f, s, payoff::digital_call(k));
			}
			else {
				return o.value(f, s, payoff::digital_put(-k));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return XLL_NAN;
}

AddIn xai_option_digital_call_value(
	Function(XLL_DOUBLE, "xll_option_digital_call_value", "OPTION.DIGITAL_CALL.VALUE")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the digital call strike."),
		})
	.FunctionHelp("Return the digital call option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_call_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_value(m, f, s, k);
}
AddIn xai_option_digital_put_value(
	Function(XLL_DOUBLE, "xll_option_digital_put_value", "OPTION.DIGITAL_PUT.VALUE")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the strike."),
		})
	.FunctionHelp("Return the digital put option value.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_put_value(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_value(m, f, s, -k);
}

AddIn xai_option_digital_delta(
	Function(XLL_DOUBLE, "xll_option_digital_delta", "OPTION.DIGITAL.DELTA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
	.FunctionHelp("Return the digital call (k > 0) or put (k < 0) option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			if (k >= 0) {
				return o.delta(f, s, payoff::digital_call(k));
			}
			else {
				return o.delta(f, s, payoff::digital_put(-k));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return XLL_NAN;
}
AddIn xai_option_digital_call_delta(
	Function(XLL_DOUBLE, "xll_option_digital_call_delta", "OPTION.DIGITAL_CALL.DELTA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
	.FunctionHelp("Return the digital call option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_call_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_delta(m, f, s, k);
}
AddIn xai_option_digital_put_delta(
	Function(XLL_DOUBLE, "xll_option_digital_put_delta", "OPTION.DIGITAL_PUT.DELTA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the digital put strike."),
		})
	.FunctionHelp("Return the digital put option delta.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_put_delta(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_delta(m, f, s, -k);
}

AddIn xai_option_digital_gamma(
	Function(XLL_DOUBLE, "xll_option_digital_gamma", "OPTION.DIGITAL.GAMMA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
	.FunctionHelp("Return the digital call (k > 0) or put (k < 0) option gamma.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_gamma(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			if (k >= 0) {
				return o.gamma(f, s, payoff::digital_call(k));
			}
			else {
				return o.gamma(f, s, payoff::digital_put(-k));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return XLL_NAN;
}
AddIn xai_option_digital_call_gamma(
	Function(XLL_DOUBLE, "xll_option_digital_call_gamma", "OPTION.DIGITAL_CALL.GAMMA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
	.FunctionHelp("Return the digital call option gamma.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_call_gamma(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_gamma(m, f, s, k);
}
AddIn xai_option_digital_put_gamma(
	Function(XLL_DOUBLE, "xll_option_digital_put_gamma", "OPTION.DIGITAL_PUT.GAMMA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
	.FunctionHelp("Return the digital put option gamma.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_put_gamma(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_gamma(m, f, s, -k);
}

AddIn xai_option_digital_vega(
	Function(XLL_DOUBLE, "xll_option_digital_vega", "OPTION.DIGITAL.VEGA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
	.FunctionHelp("Return the digital call (k > 0) or put (k < 0) option vega.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_vega(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	try {
		handle<variate_base<>> m_(m);

		if (m_) {
			option o(*m_.ptr());

			if (k >= 0) {
				return o.vega(f, s, payoff::digital_call(k));
			}
			else {
				return o.vega(f, s, payoff::digital_put(-k));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return XLL_NAN;
}
AddIn xai_option_digital_call_vega(
	Function(XLL_DOUBLE, "xll_option_digital_call_vega", "OPTION.DIGITAL_CALL.VEGA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
		.FunctionHelp("Return the digital call option vega.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_call_vega(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_vega(m, f, s, k);
}
AddIn xai_option_digital_put_vega(
	Function(XLL_DOUBLE, "xll_option_digital_put_vega", "OPTION.DIGITAL_PUT.VEGA")
	.Arguments({
		Arg(XLL_HANDLE, "m", "is a handle to a variate."),
		Arg(XLL_DOUBLE, "f", "is the forward."),
		Arg(XLL_DOUBLE, "s", "is the vol."),
		Arg(XLL_DOUBLE, "k", "is the put strike."),
		})
	.FunctionHelp("Return the digital put option vega.")
	.Category("Option")
	.HelpTopic("https://keithalewis.github.io/math/op.html")
);
double WINAPI xll_option_digital_put_vega(HANDLEX m, double f, double s, double k)
{
#pragma XLLEXPORT
	return xll_option_digital_vega(m, f, s, -k);
}
