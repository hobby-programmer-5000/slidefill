[English](README.md) | **한국어**

# 실제 Excel에서 실행하는 VBA 비교 실험

기존 SlideFill의 Python/Fraction 계산을 유지하면서, 설치된 데스크톱 Excel이 실행하는 자체 VBA 모듈과 비교합니다. 앱의 기본 계산 경로는 바꾸지 않습니다. [SlideFillOracle.bas](SlideFillOracle.bas)에는 파일·네트워크 접근이나 통합 문서 자동 실행 이벤트가 없습니다.

비교할 경로는 세 가지입니다.

1. Excel 시트 수식이 지정한 셀에 대해 VBA `SFCalc` 사용자 함수를 호출합니다.
2. 비교 실행기가 Excel `Application.Run`으로 같은 VBA 함수를 호출하고 결과를 받습니다.
3. 기존 Python 산술 코드가 같은 명시적 연산을 별도로 계산합니다.

앞의 두 경로는 VBA 구현과 Excel 런타임을 공유합니다. 일치는 연결·전달의 일관성을 보여 주며, 알고리즘을 독립적으로 증명하지는 않습니다. Python과 별도로 정한 유리수 기대값으로 공통 숫자 가정을 점검합니다. `Evaluate`나 `WorksheetFunction`만 호출한 결과는 이 VBA 모듈의 실행 증거가 아닙니다.

## 준비와 최초 가져오기

VBA가 설치된 Windows 데스크톱 Excel과 사용자의 대화형 세션이 필요합니다. 앱에 Microsoft Excel/VBA 바이너리를 묶지 않습니다. 제공하는 매크로 통합 문서와 자체 모듈은 현재 Office 정책에서 허용되어야 합니다. 이 실험은 보안 센터, 신뢰할 수 있는 위치, 서명 키, 레지스트리 설정을 변경하지 않습니다.

시험한 환경에서는 VBA 프로젝트에 대한 프로그램 접근이 불가능했습니다. 실행 가능한 준비·비교 도구를 확인한 뒤 아래 수동 단계 한 번을 진행할 수 있습니다.

1. 생성한 **시험용 통합 문서**를 Excel에서 엽니다. 원본 입력 파일은 보존합니다.
2. **Alt+F11**로 Visual Basic 편집기를 엽니다. **파일 → 파일 가져오기…**에서 이 폴더의 `SlideFillOracle.bas`를 선택하고 Excel로 돌아옵니다.
3. 시험 문서를 기존 매크로 사용 통합 문서인 `.xlsm` 형식으로 저장합니다. 모듈의 `Attribute` 줄을 일반 코드 창에 붙여 넣지 마세요. 파일 가져오기가 이 줄을 처리합니다.
4. Office 정책이 자체 시험 매크로를 막으면 거부 결과를 기록합니다. 강제로 시험하려고 모든 매크로를 허용하거나 프로젝트 접근 설정을 바꾸지 않습니다.
5. 검토한 파일로 제공된 비교 절차를 실행합니다. 가져오기와 실제 실행 전에는 두 VBA 경로 모두 **미실행(NOT RUN)**입니다. 정적 소스 검사, `#NAME?`가 보이는 수식, Excel 워크시트 함수 호출을 VBA 실행 통과로 기록하지 않습니다.

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

기존 [Python 산술](../../facts.py), [core](../../core.py), 수치 보호 장치와 테스트를 유지합니다. 실제 통합 검사에는 시험 자료·모듈 해시, 선택 셀, Excel 버전/빌드, 두 VBA 호출 경로, Python 기대값, 원본 해시 보존을 기록해야 합니다. 실험 준비만으로 새 빌드·릴리스·기본 계산 경로 교체를 뜻하지 않습니다.

## 1차 출처와 라이선스

Microsoft 문서의 [Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run), [VBA Round](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/round-function), [Double](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/double-data-type), [Decimal](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/decimal-data-type), [Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2), [Text](https://learn.microsoft.com/en-us/office/vba/api/excel.range.text), [매크로 보안](https://support.microsoft.com/en-us/excel/change-macro-security-settings-in-excel)을 확인했습니다. [VBA는 Office에 통합](https://learn.microsoft.com/en-us/visualstudio/vsto/vba-and-office-solutions-in-visual-studio-compared?view=visualstudio)되어 있습니다. 소스 모듈을 앱에 넣는 것만으로 독립 실행 엔진이나 Microsoft 바이너리 재배포 권한이 생기지 않습니다.

직접 작성한 모듈·도구는 기존 프로젝트의 [MIT 라이선스](../../LICENSE)를 따르며 고지문을 보존합니다. 타사 매크로 소스, 출처가 불명확한 `vbaProject.bin`, 새 라이브러리, Microsoft 런타임 바이너리를 포함하지 않습니다. 설치된 Office/VBA에는 자체 라이선스와 정책이 적용됩니다. Office의 구체적인 재배포 권리는 확인하지 않았으며, 새 라이선스 체계도 적용하지 않습니다.
