// xlloption.cpp - Sample xll project.
#include <cmath>
//Uncomment to build for versions of Excel prior to 2007
//#define XLOPERX XLOPER
#include "../../keithalewis/fmsoption/fms_option.h"
#include "../../keithalewis/fmsoption/fms_normal.h"
#include "xll/xll/xll.h"

using namespace xll;

struct bm : public fms::option_nvi<double, double, double> {};
struct om : public fms::option_model<fms::variate::normal<double>, double, double, double> {};
//using om = fms::option_model<fms::variate::normal<double>, double, double, double>;

AddIn xai_option_model_normal(
	Function(XLL_HANDLE, "xll_option_model_normal", "OPTION.MODEL.NORMAL")
	.Category("Option")
	.FunctionHelp("Return handle to normal option model.")
	.Uncalced()
);
HANDLEX WINAPI xll_option_model_normal()
{
#pragma XLLEXPORT
	//handle<fms::option_model<fms::variate::normal<>>> h(new fms::option_model<fms::variate::normal<>>{});
	//handle<fms::option_nvi<>> h(new fms::option_model<fms::variate::normal<>>{});
	//handle<foo> h(new foo{});

	//fms::option_nvi<>* pm = new fms::option_model<fms::variate::normal<>>();
	//handle<fms::option_nvi<>> h(pm);
	//using ob = fms::option_nvi<double, double, double>;
	//using om = fms::option_model<fms::variate::normal<double>, double, double, double>;

	handle<om> h(new om{});

	return h.get();
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
	handle<om> m_(m);

	return m_->value(f, s, k);
}
