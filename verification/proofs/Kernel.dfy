// Original project code, MIT. Dafny 4.11.0. No assume/axiom/extern/verify-false.
module VerifiedKernel {
  ghost function SumSpec(xs: seq<int>): int
    decreases |xs|
  {
    if |xs| == 0 then 0 else SumSpec(xs[..|xs|-1]) + xs[|xs|-1]
  }

  method Sum(xs: seq<int>) returns (total: int)
    ensures total == SumSpec(xs)
  {
    total := 0;
    var i := 0;
    while i < |xs|
      invariant 0 <= i <= |xs|
      invariant total == SumSpec(xs[..i])
      decreases |xs| - i
    {
      assert xs[..i+1][..i] == xs[..i];
      total := total + xs[i];
      i := i + 1;
    }
    assert xs[..i] == xs;
  }

  method Same(xs: seq<int>, ys: seq<int>) returns (ok: bool)
    ensures ok <==> xs == ys
  {
    if |xs| != |ys| { return false; }
    var i := 0;
    while i < |xs|
      invariant |xs| == |ys|
      invariant 0 <= i <= |xs|
      invariant forall j :: 0 <= j < i ==> xs[j] == ys[j]
      decreases |xs| - i
    {
      if xs[i] != ys[i] { return false; }
      i := i + 1;
    }
    assert xs == ys;
    return true;
  }

  method SumBound(deltas: seq<int>, errors: seq<int>) returns (total: int, budget: int)
    requires |deltas| == |errors|
    requires forall i :: 0 <= i < |deltas| ==> 0 <= errors[i] && -errors[i] <= deltas[i] <= errors[i]
    ensures total == SumSpec(deltas) && budget == SumSpec(errors)
    ensures 0 <= budget && -budget <= total <= budget
  {
    total := 0;
    budget := 0;
    var i := 0;
    while i < |deltas|
      invariant |deltas| == |errors|
      invariant 0 <= i <= |deltas|
      invariant total == SumSpec(deltas[..i]) && budget == SumSpec(errors[..i])
      invariant 0 <= budget && -budget <= total <= budget
      decreases |deltas| - i
    {
      assert deltas[..i+1][..i] == deltas[..i];
      assert errors[..i+1][..i] == errors[..i];
      total := total + deltas[i];
      budget := budget + errors[i];
      i := i + 1;
    }
    assert deltas[..i] == deltas && errors[..i] == errors;
  }

  method Within(a: int, b: int, epsilon: nat) returns (ok: bool)
    ensures ok <==> -(epsilon as int) <= a - b <= epsilon
  {
    return -(epsilon as int) <= a - b <= epsilon;
  }
  method RationalWithin(an: int, ad: int, bn: int, bd: int, en: nat, ed: int) returns (ok: bool)
    requires ad > 0 && bd > 0 && ed > 0
    ensures ok <==> -(en as real)/(ed as real) <= (an as real)/(ad as real) - (bn as real)/(bd as real) <= (en as real)/(ed as real)
  {
    var left := an * bd * ed - bn * ad * ed;
    var limit := en * ad * bd;
    ghost var scale: real := (ad as real) * (bd as real) * (ed as real);
    assert scale > 0.0;
    assert scale * ((an as real)/(ad as real) - (bn as real)/(bd as real)) == left as real;
    assert scale * ((en as real)/(ed as real)) == limit as real;
    ok := -limit <= left <= limit;
  }
  method RoundScaled(n: int, d: int, scale: int, quotient: int) returns (rounded: int)
    requires d > 0 && scale > 0
    requires quotient*d <= (if n < 0 then -n else n)*scale < (quotient+1)*d
    ensures n >= 0 ==> (2*rounded-1)*d <= 2*n*scale < (2*rounded+1)*d
    ensures n < 0 ==> (2*rounded-1)*d < 2*n*scale <= (2*rounded+1)*d
  {
    var magnitude := if n < 0 then -n else n;
    rounded := quotient;
    if 2*(magnitude*scale-quotient*d) >= d { rounded := rounded+1; }
    if n < 0 { rounded := -rounded; }
  }
}
