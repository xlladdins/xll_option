// xll_variate.cpp - random variates
#include "xll/xll/xll.h"
#include "../../keithalewis/fmsoption/fms_variate.h"
#include "../../keithalewis/fmsoption/fms_variate_normal.h"

using namespace fms;
using namespace xll;

AddIn xai_variate_normal(
	Function(XLL_HANDLE, "xll_variate_normal", "VARIATE.NORMAL")
	.Args({
		Arg(XLL_DOUBLE, "mu", "is the mean. Default is 0."),
		Arg(XLL_DOUBLE, "sigma", "is the standard deviation. Default is 1.")
		})
	.Category("Option")
	.FunctionHelp("Return handle to normal variate.")
	.Uncalced()
);
HANDLEX WINAPI xll_variate_normal(double mu, double sigma)
{
#pragma XLLEXPORT
	handle<variate_base<>> m(new variate_model(variate::normal(mu, sigma)));

	return m.get();
}

static AddIn xai_variate_cdf(
	Function(XLL_DOUBLE, "xll_variate_cdf", "VARIATE.CDF")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to the variate"),
		Arg(XLL_DOUBLE, "s", "is the value"),
		Arg(XLL_WORD, "n", "is the derivaive")
		}
	)
	.FunctionHelp("Return n-th derivative of cdf at s.")
	.Category("XLL")
);
double WINAPI xll_variate_cdf(HANDLEX m, double x, WORD n)
{
#pragma XLLEXPORT
	handle<variate_base<>> m_(m);

	if (m_) {
		return m_->cdf(x, n);
	}

	return std::numeric_limits<double>::quiet_NaN();
}

static AddIn xai_variate_cumulant(
	Function(XLL_DOUBLE, "xll_variate_cumulant", "VARIATE.CUMULANT")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to the variate"),
		Arg(XLL_DOUBLE, "s", "is the value"),
		Arg(XLL_WORD, "n", "is the derivaive")
		}
	)
	.FunctionHelp("Return n-th derivative of cumulant at s.")
	.Category("XLL")
);
double WINAPI xll_variate_cumulant(HANDLEX m, double x, WORD n)
{
#pragma XLLEXPORT
	handle<variate_base<>> m_(m);

	if (m_) {
		return m_->cumulant(x, n);
	}

	return std::numeric_limits<double>::quiet_NaN();
}
