// xll_variate_discrete.cpp - Excel add-in for discrete variates
#include "xll/xll/xll.h"
#include "fmsoption/fms_variate_discrete.h"
#include "fmsoption/fms_variate.h"

using namespace fms;
using namespace xll;

AddIn xai_variate_discrete(
	Function(XLL_HANDLE, "xll_variate_discrete", "VARIATE.DISCRETE")
	.Args({
		Arg(XLL_FP, "x", "are the values of the discrete random variable."),
		Arg(XLL_FP, "p", "are the probabilities of the values")
		})
	.Category("Option")
	.FunctionHelp("Return handle to the discrete variate.")
	.Uncalced()
);
HANDLEX WINAPI xll_variate_discrete(const _FPX* px, const _FPX* pp)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		ensure(size(*px) == size(*pp));
		handle<variate_base<>> m(new variate_model(variate::discrete(size(*px), px->array, pp->array)));
		if (m) {
			h = m.get();
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

