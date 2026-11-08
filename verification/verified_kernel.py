# Generated from verification/proofs/Kernel.dfy by Dafny4.11.0.
# MIT, original project code. Rebuild with proofs/build.py; do not edit.
class default__:

    def __init__(self):
        pass

    @staticmethod
    def Sum(xs):
        total: int = int(0)
        total = 0
        d_0_i_: int
        d_0_i_ = 0
        while d_0_i_ < len(xs):
            total = total + xs[d_0_i_]
            d_0_i_ = d_0_i_ + 1
        return total

    @staticmethod
    def Same(xs, ys):
        ok: bool = False
        if len(xs) != len(ys):
            ok = False
            return ok
        d_0_i_: int
        d_0_i_ = 0
        while d_0_i_ < len(xs):
            if xs[d_0_i_] != ys[d_0_i_]:
                ok = False
                return ok
            d_0_i_ = d_0_i_ + 1
        ok = True
        return ok
        return ok

    @staticmethod
    def SumBound(deltas, errors):
        total: int = int(0)
        budget: int = int(0)
        total = 0
        budget = 0
        d_0_i_: int
        d_0_i_ = 0
        while d_0_i_ < len(deltas):
            total = total + deltas[d_0_i_]
            budget = budget + errors[d_0_i_]
            d_0_i_ = d_0_i_ + 1
        return (total, budget)

    @staticmethod
    def Within(a, b, epsilon):
        ok: bool = False
        ok = 0 - epsilon <= a - b and a - b <= epsilon
        return ok
        return ok

    @staticmethod
    def RationalWithin(an, ad, bn, bd, en, ed):
        ok: bool = False
        d_0_left_: int
        d_0_left_ = an * bd * ed - bn * ad * ed
        d_1_limit_: int
        d_1_limit_ = en * ad * bd
        ok = 0 - d_1_limit_ <= d_0_left_ and d_0_left_ <= d_1_limit_
        return ok

    @staticmethod
    def RoundScaled(n, d, scale, quotient):
        rounded: int = int(0)
        d_0_magnitude_: int
        if n < 0:
            d_0_magnitude_ = 0 - n
        elif True:
            d_0_magnitude_ = n
        rounded = quotient
        if 2 * (d_0_magnitude_ * scale - quotient * d) >= d:
            rounded = rounded + 1
        if n < 0:
            rounded = 0 - rounded
        return rounded
