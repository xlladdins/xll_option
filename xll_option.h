// xlloption.h - Excel add-in for fmsoption
#pragma once
#include "xll/xll/xll.h"
#include "../../keithalewis/fmsoption/fms_option.h"
#include "../../keithalewis/fmsoption/fms_variate.h"

namespace xll {

	// Option parameterized by model and pointer to member function
	template</*class M,*/ class... Args>
	inline double xll_option(HANDLEX m, double (fms::option<fms::variate_base<>>::*pmf)(Args... args) const, Args... args)
	{
		try {
			handle<fms::variate_base<>> m_(m);

			if (m_) {
				fms::option o(*m_.ptr());
				return (o.*pmf)(args...);
				/*
				M* pm = m_.as<M>();
				if (pm) {
					const auto& o = fms::option(*pm);
					return (o.*pmf)(args...);
				}
				*/
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

	template<class M, class X = M::type, class S = M::ctype>
	using model = fms::variate_model<M, X, S>;

	template<class M>
	inline double xll_option_value(HANDLEX m, double f, double s, double k)
	{
		return xll_option<model<M>>(m, &fms::option<model<M>>::value, f, s, k);
	}
	template<class M>
	inline double xll_option_delta(HANDLEX m, double f, double s, double k)
	{
		return xll_option<model<M>>(m, &fms::option<model<M>>::delta, f, s, k);
	}
	template<class M>
	inline double xll_option_gamma(HANDLEX m, double f, double s, double k)
	{
		return xll_option<model<M>>(m, &fms::option<model<M>>::gamma, f, s, k);
	}
	template<class M>
	inline double xll_option_vega(HANDLEX m, double f, double s, double k)
	{
		return xll_option<model<M>>(m, &fms::option<model<M>>::vega, f, s, k);
	}
	
	template<class M>
	inline double xll_option_implied(HANDLEX m, double f, double s, double k, double s0, size_t n, double eps)
	{
		return xll_option<model<M>>(m, &fms::option<model<M>>::implied, f, s, k, s0, n, eps);
	}

}

