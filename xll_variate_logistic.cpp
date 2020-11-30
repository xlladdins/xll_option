// xll_variate_logistic.cpp - Excel add-in for logistic variates
#include "xll/xll/xll.h"
#include "fmsoption/fms_variate_logistic.h"
#include "fmsoption/fms_variate_handle.h"

using namespace fms;
using namespace xll;

// int breakme = [&]() { return _crtBreakAlloc = 620; }();

AddIn xai_variate_logistic(
	Function(XLL_HANDLE, "xll_variate_logistic", "VARIATE.LOGISTIC")
	.Args({
	})
	.Uncalced()
	.FunctionHelp("Return handle to logistic variate.")
	.Category("Option")
);
HANDLEX WINAPI xll_variate_logistic()
{
#pragma XLLEXPORT
	handle<variate_base<>> m(new variate_handle(variate::logistic()));

	return m.get();
}

