// xll_option.h
#pragma once
#include "fms_variate/fms_variate.h"
#include "xll/xll/xll.h"

namespace fms {

	// NVI base class for variates
	template<class X = double, class S = double>
	struct variate_base_impl {
		typedef typename X xtype;
		typedef typename S stype;

		variate_base_impl()
		{ }
		variate_base_impl(const variate_base_impl&) = delete;
		variate_base_impl& operator=(const variate_base_impl&) = delete;
		virtual ~variate_base_impl()
		{ }

		X pdf(X x, S s = 0, size_t n = 0) const
		{
			return cdf_(x, s, n - 1);
		}
		// transformed cumulative distribution function and derivatives
		X cdf(X x, S s = 0, size_t n = 0) const
		{
			return cdf_(x, s, n);
		}
		// (d/ds)^n log E[exp(sX)]
		S cumulant(S s, size_t n = 0) const
		{
			return cumulant_(s, n);
		}
		X edf(S s, X x = 0) const
		{
			return edf_(s, x);
		}
	private:
		virtual X cdf_(X x, S s, size_t n) const = 0;
		virtual S cumulant_(S s, size_t n) const = 0;
		virtual X edf_(S s, X x) const = 0;
	};

	template<class X = double, class S = X>
	using variate_base = variate_model<variate_base_impl<X, S>>;

	// implement for a specific model and make a copy
	template<class M, class X = M::xtype, class S = M::stype>
	class variate_handle : public variate_base<X, S>
	{
		M m;
	public:
		variate_handle(const M& m)
			: m(m)
		{ }
		variate_handle(const variate_handle&) = default;
		variate_handle& operator=(const variate_handle&) = default;
		~variate_handle()
		{ }

		X cdf_(X x, S s = 0, size_t n = 0) const override
		{
			return m.cdf(x, s, n);
		}
		S cumulant_(S s, size_t n = 0) const override
		{
			return m.cumulant(s, n);
		}
		X edf_(S s, X x) const override
		{
			return m.edf(s, x);
		}
	};

}