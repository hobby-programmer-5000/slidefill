[English](README.md) | **한국어**

# 실제 Excel에서 실행하는 VBA 비교 실험

기존 SlideFill의 Python/Fraction 계산을 유지하면서, 설치된 데스크톱 Excel이 실행하는 자체 VBA 모듈과 비교합니다. 앱의 기본 계산 경로는 바꾸지 않습니다. [SlideFillOracle.bas](SlideFillOracle.bas)에는 파일·네트워크 접근이나 통합 문서 자동 실행 이벤트가 없습니다.

앱의 **Excel VBA 검산**에서도 지원하는 보고서 집계에 같은 `SFCalc` 모듈을 사용합니다. [내장 통합 문서](../../assets/slidefill-vba.xlsm)에는 빈 시트 두 장과 실험에 사용한 VBA 프로젝트가 들어 있습니다. 재현하려면 검증을 마친 단일 모듈 통합 문서를 입력으로 `python tools/vba/make_runtime_engine.py TESTED.xlsm NEW.xlsm`을 실행합니다. 이 도구는 VBA 프로젝트 해시를 확인한 뒤 원본 시트 데이터를 복사하지 않고 새 파일을 만들며, 기존 파일을 덮어쓰지 않습니다.

비교할 경로는 세 가지입니다.

1. Excel 시트 수식이 지정한 셀에 대해 VBA `SFCalc` 사용자 함수를 호출합니다.
2. 비교 실행기가 Excel `Application.Run`으로 같은 VBA 함수를 호출하고 결과를 받습니다.
3. 기존 Python 산술 코드가 같은 명시적 연산을 별도로 계산합니다.

앞의 두 경로는 VBA 구현과 Excel 런타임을 공유합니다. 일치는 연결·전달의 일관성을 보여 주며, 알고리즘을 독립적으로 증명하지는 않습니다. Python과 별도로 정한 유리수 기대값으로 공통 숫자 가정을 점검합니다. `Evaluate`나 `WorksheetFunction`만 호출한 결과는 이 VBA 모듈의 실행 증거가 아닙니다.

## 준비와 최초 가져오기

VBA가 설치된 Windows 데스크톱 Excel과 사용자의 대화형 세션이 필요합니다. 앱에 Microsoft Excel/VBA 바이너리를 묶지 않습니다. 제공하는 매크로 통합 문서와 자체 모듈은 현재 Office 정책에서 허용되어야 합니다. 이 실험은 보안 센터, 신뢰할 수 있는 위치, 서명 키, 레지스트리 설정을 변경하지 않습니다.

시험한 환경에서는 VBA 프로젝트에 대한 프로그램 접근이 불가능했습니다. 실행 가능한 준비·비교 도구를 확인한 뒤 아래 수동 단계 한 번을 진행할 수 있습니다.

1. 생성한 **시험용 통합 문서**를 Excel에서 엽니다. 원본 입력 파일은 보존합니다.
2. **Alt+F11**로 Visual Basic 편집기를 엽니다. 왼쪽 프로젝트 탐색기에서 **VBAProject (시험 파일 이름)**을 먼저 선택합니다. **File → Import File…(파일 → 파일 가져오기…)**에서 이 폴더의 `SlideFillOracle.bas`를 선택합니다. 같은 프로젝트 아래 **Modules → SlideFillOracle**이 생긴 것을 확인한 뒤 Excel로 돌아옵니다. `Sheet1`, `Sheet2`, `ThisWorkbook`만 보이면 아직 기본 객체만 있는 상태입니다.
3. 시험 문서를 기존 매크로 사용 통합 문서인 `.xlsm` 형식으로 저장합니다. 모듈의 `Attribute` 줄을 일반 코드 창에 붙여 넣지 마세요. 파일 가져오기가 이 줄을 처리합니다.
4. Office 정책이 자체 시험 매크로를 막으면 거부 결과를 기록합니다. 강제로 시험하려고 모든 매크로를 허용하거나 프로젝트 접근 설정을 바꾸지 않습니다.
5. 검토한 파일로 제공된 비교 절차를 실행합니다. 가져오기와 실제 실행 전에는 두 VBA 경로 모두 **미실행(NOT RUN)**입니다. 정적 소스 검사, `#NAME?`가 보이는 수식, Excel 워크시트 함수 호출을 VBA 실행 통과로 기록하지 않습니다.

시험용 통합 문서마다 `SlideFillOracle`을 한 번만 가져옵니다. 두 번째 가져오기로 `SlideFillOracle1`이 생겼다면 시험 파일에서 그 중복 모듈만 제거합니다. 이번 실행에서는 중복을 제거한 뒤 `SFCalc` 이름 충돌이 해소됐습니다.

예를 들어 영문 Excel의 `=SFCalc("mean",B2:B4)`는 세 셀을 대상으로 모듈을 호출합니다. 지역 설정에 따라 수식 구분자가 다를 수 있습니다. 실행기는 `'Engine.xlsm'!SlideFillOracle.SFCalc`처럼 통합 문서까지 지정한 이름으로 검증된 모듈을 호출할 수 있습니다. 현재 활성 문서에 의존하지 말고 실제 모듈 파일과 입력의 식별 정보를 확인해야 합니다.

## 실행 명령

저장소 루트에서 기존 프로젝트 가상환경을 사용합니다. **새 실행 폴더**를 선택하고 이전 파일·검사 기록을 보존합니다. 아래 예시는 이미 있을 수 있는 `run-01`을 보존하도록 `run-02`를 사용합니다.

```powershell
$spikeRun = '.cache/agent-work/vba-spike/run-02'
.venv/Scripts/python.exe tools/vba/compare.py prepare $spikeRun
powershell -NoProfile -File tools/vba/host.ps1 -Mode Build -InputWorkbook "$spikeRun/inputs.xlsx" -ModulePath tools/vba/SlideFillOracle.bas -EngineWorkbook "$spikeRun/engine.xlsm" -OutputJson "$spikeRun/build.json"
```

`build.json`을 읽습니다. 프로젝트 접근 제한으로 자동 가져오기가 막혔으면 위의 일반 VBE 파일 가져오기 절차로 해당 실행 폴더의 `engine.xlsm`에 모듈을 가져온 뒤 저장하고 닫습니다. 차단된 Build 결과는 VBA 실행 결과가 아닙니다. Build를 통과시키려고 프로젝트 접근·매크로 보안 설정을 변경하지 않습니다.

```powershell
powershell -NoProfile -File tools/vba/host.ps1 -Mode Run -EngineWorkbook "$spikeRun/engine.xlsm" -AppInputJson "$spikeRun/app-input.json" -OutputJson "$spikeRun/native-results.json"
.venv/Scripts/python.exe tools/vba/compare.py summarize "$spikeRun/app-input.json" "$spikeRun/native-results.json" "$spikeRun/comparison.json"
```

준비한 manifest에는21개 사례가 있습니다. `native-results.json`과 `comparison.json`에서 실제 실행 건수를 확인하고 MATCH, DIFFERENT, 차단/계산 불가, 오류를 구분합니다. 선언한 비교 계약상 이진수·반올림 차이가 예상되더라도 차이를 일치로 바꾸지 않습니다. 원본·모듈·실행 통합 문서·입력 해시를 결과와 함께 보존합니다. 매크로를 가져오거나 허용하거나 실제 실행하지 못했다면 파일을 열었다는 이유로 성공을 추정하지 말고 **미실행(NOT RUN)**으로 기록합니다.

Run은 VBA 호출 전에 원본 파일과 실제 엔진 셀의 형식·Double 비트를 대조합니다. 선언한 해시만으로 입력이 같다고 판단하지 않습니다. Build는 숫자처럼 보이는 텍스트와 표시 형식 문자열을 명시적으로 보존합니다. 예전 시험 파일은 Excel `Value2` 복사 중 이러한 문자열이 숫자로 바뀔 수 있었습니다. 모듈을 가져와 저장한 뒤 `repair_fixture.py`로 새 사본을 만들면 원본 파일과 내장 VBA 바이너리를 그대로 보존하면서 시험 데이터 형식만 복구할 수 있습니다.

```powershell
.venv/Scripts/python.exe tools/vba/repair_fixture.py "$spikeRun/inputs.xlsx" "$spikeRun/engine-final.xlsm" "$spikeRun/engine-typed.xlsm"
```

이후 새 사본을 `-EngineWorkbook`으로 사용합니다. VBA 프로젝트가 없거나 출력 파일이 이미 있으면 거부합니다. 복구 도구의 패키지 검사는 실행 불가능한 표식 바이트를 사용하며 VBA 컴파일·실행을 입증하지 않습니다.

## 실제 실행 결과

Excel 16.0(빌드 20326)에서 21개 사례의 원본 통합 문서, 엔진 사본, 앱 입력을 비교해 피연산자 형식과 숫자의 Double 비트를 확인했습니다. 시트의 VBA 함수와 `Application.Run` 결과는 **21건 중 21건** 일치했습니다. 16건은 Double 비트가 같았고 5건은 오류 코드가 같았습니다. Python 결과를 가장 가까운 Double로 바꾸거나 예상 오류 코드를 비교하면 **21건 중 17건**이 일치했습니다.

| 사례 | 두 VBA 경로의 결과 | Python의 정확값 |
|---|---:|---:|
| `decimal_sum` — `0.1 + 0.2`를 Double로 합산 | `0.30000000000000004` | `3/10` |
| `cancellation` — `10^16 + 1 − 10^16` | `0` | `1` |
| `tiny_growth` — `100`에서 `100.00000001`로 증가 | `9.9999937219763522e-11` | `1/10000000000` |
| `banker_tie_probe` — `2.5`를 정수로 반올림 | `2` | `3` (절반에서 0과 멀어지는 방향) |

`decimal_variant`에서 `CDec`으로 `0.1 + 0.2`를 합산한 결과는 `3/10`에 가장 가까운 Double이었습니다. `large_literal`에서는 두 VBA 경로 모두 `9007199254740992`를 반환했습니다. Double 경계에서는 일치했지만 정확한 입력 `9007199254740993`보다 1 작았습니다. Excel의 `Range.Text`가 반환한 표시 문자열은 `9007199254740990`이었습니다.

이 결과는 개발용 비교 실행기의 가상 자료 21개에 해당합니다. SlideFill의 보고서 생성은 Python/Fraction으로 계산합니다. 실행 기록은 `.cache/agent-work/vba-spike/run-02/`의 `native-results-after-dedup.json`과 `comparison-after-dedup.json`에 있습니다.

## 고정한 연산 계약

`comparison.json`은 두 VBA 경로의 일치, 선언한 Double 출력 경계에서의 일치, 정확한 차이인 `native_binary_minus_python_exact`를 따로 기록합니다. 큰 정수는 Double 변환에서 1이 사라져도 두 VBA 경로가 같을 수 있으므로 정확한 차이를 함께 봅니다. 진단 결과로 Python 계산값을 조용히 바꾸지 않습니다.

`SFCalc(operation, values, [digits])`는 연속된 Excel 범위, 2차원 Variant 배열 또는 단일 값을 받습니다. 범위·배열은 **100,000셀 이하**이고 행 우선 순서로 읽습니다. 범위는 `Value2`로 읽습니다. 실제 빈 셀은 제외하고, 빈 문자열을 포함한 텍스트·Boolean·Date 형식 값·Excel 오류·객체·유한하지 않은 숫자는 거부합니다. `Value2`가 반환한 날짜 일련번호는 숫자이므로 호출자가 실제 수치 셀을 명시적으로 골라야 합니다.

| 연산 | 결과와 조건 |
|---|---|
| `sum` | 남긴 수치의 Double 합. 전부 비었으면0. |
| `count` | 남긴 수치 개수를 Double로 반환. 전부 비었으면0. |
| `mean` | Double 합/개수. 전부 비었으면 0 나눗셈 상태. |
| `growth` | 남긴 수치가 정확히 두 개이며 **현재값, 기준값** 순서. `(현재-기준)/Abs(기준)`. 기준0은 계산 불가. |
| `points` | **현재값, 기준값** 두 개의 절대 차이. 상대 증감률로 취급하지 않음. |
| `round` | 수치 한 개, 정수 자릿수−15..15. `WorksheetFunction.Round` 사용. |
| `vba_round_probe` | 수치 한 개, 정수 자릿수0..15. 은행가 반올림을 하는 VBA `Round`를 의도적으로 비교하는 진단용 연산. |
| `decimal_sum` | 제한된 `CDec` 누산 뒤 최종 Double로 반환. 진단용이며 임의 정밀도 유리수 산술은 아님. |

오류 결과는 실제 VBA `CVErr`입니다. 지원하지 않는 연산·배열 모양·입력 형식·자릿수·인수 개수는 **2015 / `#VALUE!`**, 빈 평균·기준0은 **2007 / `#DIV/0!`**, 산술 실패·Decimal 변환/범위 초과·유한하지 않은 숫자는 **2036 / `#NUM!`**입니다. COM 연결에서는 오류 Variant나 HRESULT 기반 래퍼로 보일 수 있습니다. 오류 정수를 정상 숫자로 해석하지 말고 형식과 상태를 보존해야 합니다.

기여 셀, 상태, 원시 숫자, 선언한 표시 문자열, 최종 보고서 전달을 따로 비교합니다. Excel의 `Range.Text`는 열 너비에 따라 달라질 수 있습니다. 결과를 억지로 맞추려고 ‘표시된 자릿수로 정밀도 설정’이나 원본 값·서식을 바꾸지 않습니다.

## 한계와 증거

명시적 입력에 대한 위 연산만 다룹니다. 자동 선택, 모든 수식, 날짜·달력, 임의의 사용자 매크로, 보고서 렌더링을 다시 구현하지 않습니다. Double 반복 합이나 제한된 Decimal 중간값은 Python/Fraction이 정확하게 계산하는 경우에도 정보를 잃거나 범위를 넘을 수 있습니다. Excel 숫자 셀에서 이미 사라진 자릿수는 VBA로 복원되지 않습니다. 상쇄, 반올림 경계, 음수·0 기준, 빈칸·오류, 미세한 변화를 시험 자료에 포함해야 합니다.

보고서 값은 기존 [Python 산술](../../facts.py), [core](../../core.py), 수치 보호 장치로 계산합니다. 앱 검산은 독립 검증을 통과한 표·차트의 단일 열 합계·평균·개수를 대상으로 합니다. 숫자 형태의 텍스트, 날짜, 수식, 필터나 여러 열을 쓰는 집계는 첫 검산 범위에서 제외합니다.

## 1차 출처와 라이선스

Microsoft 문서의 [Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run), [VBA Round](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/round-function), [Double](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/double-data-type), [Decimal](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/decimal-data-type), [Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2), [Text](https://learn.microsoft.com/en-us/office/vba/api/excel.range.text), [매크로 보안](https://support.microsoft.com/en-us/excel/change-macro-security-settings-in-excel)을 확인했습니다. [VBA는 Office에 통합](https://learn.microsoft.com/en-us/visualstudio/vsto/vba-and-office-solutions-in-visual-studio-compared?view=visualstudio)되어 있습니다. 소스 모듈을 앱에 넣는 것만으로 독립 실행 엔진이나 Microsoft 바이너리 재배포 권한이 생기지 않습니다.

직접 작성한 모듈과 이를 컴파일한 VBA 프로젝트, 도구는 기존 [MIT 라이선스](../../LICENSE)를 따릅니다. 설치된 Office/VBA에는 호스트의 라이선스와 정책이 적용됩니다.
