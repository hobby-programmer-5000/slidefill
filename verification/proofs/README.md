**English** | [한국어](README.ko.md)

# Verified kernel used by the app

`Kernel.dfy` contains no `assume`, axioms, extern bodies or verification-disable directives. Dafny4.11.0 reports9verified units,0errors for7source definitions, including the recursive specification/termination and these6executable contracts:

- `Sum(xs) == SumSpec(xs)` for every finite integer sequence; the loop terminates.
- `Same(xs,ys)` is true exactly when the two finite integer sequences are equal, including length and order; the loop terminates.
- `Within(a,b,epsilon)` is true exactly when `-epsilon <= a-b <= epsilon`, with nonnegative epsilon.
- `RationalWithin(an,ad,bn,bd,en,ed)` computes the integer cross product and is true exactly when `-en/ed <= an/ad-bn/bd <= en/ed` as real-rational arithmetic, with positive denominators and nonnegative `en`. Its real scale exists only in ghost proof code.

- `SumBound(deltas,errors)` returns their exact integer sums and proves `0 <= budget` and `-budget <= total <= budget`, when lengths match and every nonnegative input error bounds its signed delta. `kernel.sum_error_bounds` establishes these conditions; automatic aggregate recipes use the generated result for independently reconstructed row-decoding errors. Three wrong sum/budget/loop implementations were rejected under unchanged contracts.

- `RoundScaled(n,d,scale,quotient)` proves signed half-up integer rounding, conditional on positive `d/scale` and the checked floor-quotient bounds. `kernel.round_decimal` establishes those bounds and `numfmt._digits` uses the generated result in actual display. See [the exact interval](../../doc/binary-verification.md).

`verified_kernel.py` is the actual generated class used by `kernel.py`. The build script removes unused generated imports only, preserving the compiled class AST exactly. This integer-only subset needs no Dafny runtime in the app. `layouts._cost_sum` calls the generated sum through an exact scaled-decimal bridge. Automatic table/text comparisons call the generated equality through length-prefixed integer encodings; rational numeric guards call the generated and proved positive-denominator cross multiplication. Fraction normalization/argument extraction and the other type/conversion bridges remain tested/trusted, not formally verified. Three wrong rational numerator/bound/sign implementations were also rejected with unchanged contracts.

To reproduce as a developer, acquire the official [Dafny4.11.0 Windows release](https://github.com/dafny-lang/dafny/releases/tag/v4.11.0), verify SHA256 `3653f05a111ca21e234709ea7b25ce96083fd6c6f10484256ba54110dbc0654d`, then run:

```powershell
& .\.venv\Scripts\python.exe verification\proofs\build.py C:\tools\dafny\Dafny.exe C:\temp\new-proof-run
```

The output directory must be new. Inspect the log, exit status, exact generated hash and `receipt.json`. Do not use `--no-verify`. The shipped receipt pins the source and generated-class hashes. Wrong-sum, always-equal-length and one-sided-tolerance implementation mutations with unchanged contracts were tested and rejected by the solver. Runtime property/mutation tests additionally check the bridges and guard, not just the Dafny model.

Trusted base: Dafny frontend/backend, Boogie/Z3, the Python integer model and runtime, the AST-preserving extraction script, rational/decimal/string conversions, source/selection recipes, authentic input bytes and serialized-PPTX inspection. Finite-resource exhaustion is outside abstract termination. These contracts do not prove the whole GUI, all business intent, spreadsheet semantics, float arithmetic or Office rendering.

Licensing recommendation: keep the existing MIT license for original project/proof code. Dafny itself is [MIT](https://github.com/dafny-lang/dafny/blob/v4.11.0/LICENSE.txt); tool binaries and their dependencies stay in ignored developer storage and are not shipped. No third-party runtime or model was added to the app. See [automatic verification usage](../../doc/verification.md) and [algorithm and code map](../../doc/verification-design.md).
