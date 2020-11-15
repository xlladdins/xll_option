// xlloption.h - Excel add-in for fmsoption
#pragma once
#include "xll/xll/xll.h"
#include "../../keithalewis/fmsoption/fms_option.h"
#include "../../keithalewis/fmsoption/fms_variate.h"

namespace xll {

	template<class M>
	inline double xll_option(HANDLEX m, double (fms::option<M>::*pmf)(double,double,double) const, double f, double s, double k)
	{
		handle<fms::variate_base<>> m_(m);

		if (m_) {
			M* pm = m_.as<M>();
			if (pm) {
				const auto& o = fms::option(*pm);
				return (o.*pmf)(f, s, k);
			}
		}

		return std::numeric_limits<double>::quiet_NaN();
	}

	template<class M>
	inline double xll_option_value(HANDLEX m, double f, double s, double k)
	{
		return xll_option<M>(m, &fms::option<M>::value, f, s, k);
	}
}

