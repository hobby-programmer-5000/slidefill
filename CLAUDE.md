# SlideFill

이 프로젝트는 `small-app-harness` 스킬의 공통 규칙을 따른다. 아래는 이 프로젝트에만 해당하는 내용이다.
개인 정보(계정, 로컬 경로)는 `CLAUDE.local.md`, 진행 상태는 `PROGRESS.md`에 있다 (둘 다 gitignore).
이 파일은 공개 저장소에 올라가도 괜찮은 내용만 쓴다.

## 무엇인가

엑셀 데이터를 PowerPoint 템플릿의 `{{자리표시자}}`, 표, 차트에 채워 보고서를 만드는 Windows 포터블 앱. 정기 보고서를 만드는 사람이 쓴다.

## 명령

| 할 일 | 명령 |
|---|---|
| 테스트 | `.venv\Scripts\python.exe -X utf8 test_core.py` |
| 소스로 실행 | `.venv\Scripts\python.exe app.py` |
| 샘플 재생성 | `.venv\Scripts\python.exe -X utf8 tools\make_samples.py` |
| 빌드 (테스트, 스모크 포함) | `build.bat` → `dist\SlideFill.exe` |
| PowerPoint 실물 검증 | `powershell -File tools\verify_powerpoint.ps1 <결과.pptx>` |
| 실제 모델로 exe 검증 (Ollama 필요) | `dist\SlideFill.exe --smoke-insight <템플릿> <엑셀> <결과.json>` |

## 데이터

- 사용자 데이터: `SlideFill.json` (exe 옆, gitignore). 최근 경로와 언어만 저장한다.
- 사용자 데이터: `SlideFill.insights.json` (exe 옆, gitignore). 사용자가 고른 통찰 문장과 계산 결과의 지문.

## 이 프로젝트의 불변 조건

- 미리보기와 생성은 같은 코드 경로(`core.run`)를 쓴다. 차이는 저장 여부뿐이다.
- 값을 확정할 수 없으면 자리표시자를 그대로 남긴다. 빈칸이나 추측한 값으로 바꾸지 않는다.
- 템플릿과 엑셀 원본에는 쓰지 않는다. 출력은 항상 새 파일이고 기존 파일을 덮어쓰지 않는다.
- `core.py`, `numfmt.py`, `facts.py`, `insight.py`, `suggest.py`는 tkinter를 import하지 않는다.
- `suggest.py`의 제안은 전부 `core`로 실제 계산해 결과와 함께 보여 준다. 계산되지 않은 참조를 제안하지 않는다.
- `core.py`는 모델을 부르지 않고 `insight.py`를 import하지 않는다. 모델 문장은 사용자가 고른 것만 보고서에 들어간다.
- 네트워크 호출은 `insight.py`에만 있고, 주소는 127.0.0.1로 고정한다. 프록시를 거치지 않는다.
- 모델에게 계산을 시키지 않는다. 계산은 `facts.py`가 한다.

## 하지 말 것

- 회사 자료나 남의 템플릿을 `samples/`에 넣지 않는다. 샘플은 `tools/make_samples.py`로 직접 만든다.
- 앱 이름에 Excel, PowerPoint 같은 Microsoft 상표를 쓰지 않는다.
