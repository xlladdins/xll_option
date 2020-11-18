// xll_variate_normal.cpp - Excel add-in for normal variate
#include "xll/xll/xll.h"
#include "fmsoption/fms_variate_normal.h"
#include "fmsoption/fms_variate.h"

using namespace fms;
using namespace xll;

// int breakme = [&]() { return _crtBreakAlloc = 620; }();

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

