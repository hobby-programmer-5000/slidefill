[English](README.md) | **한국어**

# 앱이 실제 사용하는 증명 핵심

`Kernel.dfy`에는 `assume`, 공리, 외부 구현이나 검증 해제 지시문이 없습니다. Dafny 4.11.0은 소스 정의 7개에 대해 검증 단위 9개·오류 0개를 보고했습니다. 재귀 명세와 종료를 포함하며 실행 계약은 다음 여섯 가지입니다.

- 모든 유한 정수열에 대해 `Sum(xs) == SumSpec(xs)`이며 반복문이 종료합니다.
- `Same(xs,ys)`는 길이·순서를 포함한 두 유한 정수열이 같을 때만 참이며 반복문이 종료합니다.
- 음이 아닌 epsilon에 대해 `Within(a,b,epsilon)`은 `-epsilon <= a-b <= epsilon`일 때만 참입니다.
- `RationalWithin(an,ad,bn,bd,en,ed)`는 양의 분모와 음이 아닌 `en`에 대해 정수 교차곱을 계산하고, 실수·유리수 연산으로 `-en/ed <= an/ad-bn/bd <= en/ed`일 때만 참입니다. 실수 배율은 증명 전용 ghost 코드에만 있습니다.

- `SumBound(deltas,errors)`는 두 정수 목록의 정확한 합을 반환하고 `0 <= budget`·`-budget <= total <= budget`를 증명합니다. 같은 길이이고 음이 아닌 각 오차가 해당 부호 있는 차이를 감싸야 합니다. `kernel.sum_error_bounds`가 조건을 확인하며 자동 집계 검사는 독립 재구성한 행별 디코딩 차이에 생성 결과를 씁니다. 합·오차 한계·반복문이 잘못된 구현 세 가지는 동일 계약에서 거절됐습니다.

- `RoundScaled(n,d,scale,quotient)`는 양의 분모/배율과 검사한 정수 몫 범위를 전제로 부호 있는 반올림을 증명합니다. `kernel.round_decimal`이 범위를 확인하고 `numfmt._digits`가 실제 표시에 생성 결과를 씁니다. [정확한 구간](../../doc/binary-verification.ko.md)을 참고하세요.

`verified_kernel.py`는 `kernel.py`가 실제 호출하는 생성 클래스입니다. 빌드 스크립트는 사용하지 않는 생성 모듈 import만 제거하고 컴파일된 클래스의 AST를 그대로 유지합니다. 이 정수 전용 부분에는 앱용 Dafny 런타임이 필요하지 않습니다. `layouts._cost_sum`이 정확한 십진 정수 변환을 거쳐 생성된 합산을 부릅니다. 자동 표 비교는 길이를 포함한 정수 인코딩, 차트 글자 비교는 코드 포인트를 통해 목록 일치를 호출합니다. 유리수 비교는 생성되고 증명된 양의 분모 교차곱을 호출합니다. Fraction 정규화·인자 추출과 나머지 타입·변환 연결은 테스트·신뢰 경계이며 형식 증명한 것은 아닙니다. 유리수 분자·허용치·부호가 잘못된 구현 세 가지도 계약을 유지한 채 거절됐습니다.

## 개발자가 재현하는 방법

공식 [Dafny 4.11.0 Windows 릴리스](https://github.com/dafny-lang/dafny/releases/tag/v4.11.0)의 압축 파일 SHA256이 `3653f05a111ca21e234709ea7b25ce96083fd6c6f10484256ba54110dbc0654d`인지 확인하고 프로젝트 폴더에서 실행합니다.

```powershell
& .\.venv\Scripts\python.exe verification\proofs\build.py C:\tools\dafny\Dafny.exe C:\temp\new-proof-run
```

출력 폴더는 새 폴더여야 합니다. 로그·종료 상태·정확한 생성 해시·`receipt.json`을 확인하세요. `--no-verify`를 쓰지 않습니다. 함께 둔 기록은 소스와 생성 클래스의 해시를 고정합니다. 합을 틀리게 계산하는 구현, 같은 길이면 항상 같다고 하는 구현, 허용 차이의 한쪽만 검사하는 구현으로 바꾸되 계약은 유지한 변이 세 가지가 모두 솔버에서 거절됐습니다. 실행 시 속성·변이 테스트로 변환과 자동 검사도 확인합니다.

신뢰 기반은 Dafny 프런트엔드/백엔드, Boogie/Z3, Python 정수와 실행 환경, AST를 보존하는 추출 스크립트, 유리수·십진수·글자 변환, 원본 선택·계산 근거, 진짜 입력 바이트와 저장된 PPTX 검사입니다. 유한 자원 고갈은 추상적인 종료 증명의 범위 밖입니다. 이 계약은 GUI 전체·모든 업무 의미·스프레드시트 의미·부동소수점 계산·Office 렌더링을 증명하지 않습니다.

라이선스 권고는 원본 프로젝트·증명 코드의 기존 MIT 유지입니다. Dafny 자체도 [MIT](https://github.com/dafny-lang/dafny/blob/v4.11.0/LICENSE.txt)이며 도구 바이너리와 의존성은 무시되는 개발자 저장소에만 두고 배포하지 않습니다. 앱에 제3자 런타임이나 모델을 추가하지 않았습니다. [자동 검증 사용법](../../doc/verification.ko.md)과 [알고리즘·코드 구조](../../doc/verification-design.ko.md)를 참고하세요.
