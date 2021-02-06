// xll_variate_normal.cpp - Excel add-in for normal variates
#include "xll/xll/xll.h"
#include "fmsoption/fms_variate_normal.h"
#include "fmsoption/fms_variate_handle.h"

using namespace fms;
using namespace xll;

// int breakme = [&]() { return _crtBreakAlloc = 620; }();

AddIn xai_variate_normal(
	Function(XLL_HANDLE, "xll_variate_normal", "VARIATE.NORMAL")
	.Args({
		Arg(XLL_DOUBLE, "mu", "is the mean. Default is 0."),
		Arg(XLL_DOUBLE, "sigma", "is the standard deviation. Default is 1.")
	})
	.Uncalced()
	.FunctionHelp("Return handle to normal variate.")
	.Category("Option")
);
HANDLEX WINAPI xll_variate_normal(double mu, double sigma)
{
#pragma XLLEXPORT
	handle<variate_base<>> m(new variate_handle(variate::normal<>(mu, sigma)));

	return m.get();
}

