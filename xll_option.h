// xlloption.h - Excel add-in for fmsoption
#pragma once
#include "xll/xll/xll.h"
#include "../../keithalewis/fmsoption/fms_option.h"
#include "../../keithalewis/fmsoption/fms_variate.h"

namespace xll {

	// Option parameterized by model and pointer to member function
	template<class... Args>
	inline double xll_option(HANDLEX m, double (fms::option<fms::variate_base<>>::*pmf)(Args... args) const, Args... args)
	{
		try {
			handle<fms::variate_base<>> m_(m);

			if (m_) {
				fms::option o(*m_.ptr());
				return (o.*pmf)(args...);
			}
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());
		}
		catch (...) {
			XLL_ERROR("xll_option: unknown exception");
		}

		return std::numeric_limits<double>::quiet_NaN();
	}

}

