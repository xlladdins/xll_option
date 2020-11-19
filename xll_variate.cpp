// xll_variate.cpp - random variates
#include "xll/xll/xll.h"
#include "fmsoption/fms_variate.h"

using namespace fms;
using namespace xll;

static AddIn xai_variate_cdf(
	Function(XLL_DOUBLE, "xll_variate_cdf", "VARIATE.CDF")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to the variate"),
		Arg(XLL_DOUBLE, "x", "is the value"),
		Arg(XLL_DOUBLE, "s", "is the Esscher transform parameter. Default is 0."),
		Arg(XLL_WORD, "n", "is the derivative. Default is 0.")
		}
	)
	.FunctionHelp("Return s transformed n-th derivative of the cumulative distribution function at x.")
	.Category("XLL")
);
double WINAPI xll_variate_cdf(HANDLEX m, double x, double s, WORD n)
{
#pragma XLLEXPORT
	handle<variate_base<>> m_(m);

	if (m_) {
		return m_->cdf(x, s, n);
	}

	return std::numeric_limits<double>::quiet_NaN();
}

static AddIn xai_variate_pdf(
	Function(XLL_DOUBLE, "xll_variate_pdf", "VARIATE.PDF")
	.Args({
		Arg(XLL_HANDLEX, "m", "is a handle to the variate"),
		Arg(XLL_DOUBLE, "x", "is the value"),
		Arg(XLL_DOUBLE, "s", "is the Esscher transform parameter. Default is 0."),
		}
	)
	.FunctionHelp("Return s transformed probability density at x.")
	.Category("XLL")
);
double WINAPI xll_variate_pdf(HANDLEX m, double x, double s)
{
#pragma XLLEXPORT
	handle<variate_base<>> m_(m);

	if (m_) {
		return m_->cdf(x, s, 1);
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
