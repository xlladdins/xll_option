# xlloption

Excel add-in for European option pricing using [fmsoption](https://github.com/keithalewis/fmsoption).

This add-in calculates values and greeks for put and call options using any variate model.
After implementing a model in `fmsoption` create an add-in that returns a handle to a `fms::variant_base`.
See [`xll_variate_normal.cpp`](https://github.com/xlladdins/xlloption/blob/master/xll_variate_normal.cpp)
for how to do that.

The handle returned by this can be used to calculate the value and greeks of put or call options.
For example, the Excel function `=OPTION.PUT.DELTA(h, f, s, k)` returns the (forward) delta of
a put option having forward `f`, vol `s`, and strike `k` given a handle `h` for a model.
