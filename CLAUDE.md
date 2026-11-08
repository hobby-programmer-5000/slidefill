# SlideFill

이 프로젝트는 `small-app-harness` 스킬의 공통 규칙을 따른다. 아래는 이 프로젝트에만 해당하는 내용이다.
개인 정보(계정, 로컬 경로)는 `CLAUDE.local.md`, 진행 상태는 `PROGRESS.md`에 있다 (둘 다 gitignore).
이 파일은 공개 저장소에 올라가도 괜찮은 내용만 쓴다.

## 무엇인가

엑셀을 고르면 PowerPoint 보고서를 만드는 Windows 포터블 앱. 입력은 엑셀 하나가 기본이고, 회사 템플릿의 `{{자리표시자}}`, 표, 차트에 채우는 것은 선택 항목이다. 정기 보고서를 만드는 사람이 쓴다.

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
- 엑셀만으로 만들 때도 자동 선택은 `suggest.py`의 명시적 참조로 바꿔 `core.run`을 거친다. 고른 시트와 수치를 화면에 적고, 사용자가 값을 본 뒤에 만든다. 작업 템플릿은 임시 폴더에만 둔다. 넣지 않은 수치와 시트의 개수도 화면에 적는다.
- `core.py`, `numfmt.py`, `facts.py`, `insight.py`, `suggest.py`, `readers.py`, `documents.py`는 tkinter를 import하지 않는다.
- 파일 모양을 규칙으로 다 맞힐 수는 없다. 자동 선택에 의심 신호(`suggest.find_doubts`)가 있으면 보고서를 만들기 전에 사람에게 묻는다. 사람이 고르거나 "이대로 만들기"를 누른 뒤에는 다시 묻지 않는다.
- 정확도는 조정에 쓰지 않은 공개 파일로, 고치기 전에 원본과 대조해 잰다. 방법은 `tools/measure/README.md`.
- 문서와 수치의 대응은 `documents.py`가 계산한다. "같은 시기"까지만 말하고 원인이라고 쓰지 않는다. `.docx`는 zip과 XML로만 읽고 외부 엔티티를 가져오지 않는다.
- 여러 파일 비교(같은 방향인지, 비교할 수 있는지)는 `facts.py`가 계산한다. 모델은 그 결과를 읽을 뿐이고, 파일 사이의 인과는 어디에도 쓰지 않는다.
- Windows 메시지 콜백(`contents.DropTarget.handle`) 안에서는 Tk를 호출하지 않는다. 프로세스가 통째로 죽는다.
- `suggest.py`의 제안은 전부 `core`로 실제 계산해 결과와 함께 보여 준다. 계산되지 않은 참조를 제안하지 않는다.
- `core.py`는 모델을 부르지 않고 `insight.py`를 import하지 않는다. 모델 문장은 사용자가 고른 것만 보고서에 들어간다.
- 네트워크 호출은 `insight.py`에만 있고, 주소는 127.0.0.1로 고정한다. 프록시를 거치지 않는다.
- 모델에게 계산을 시키지 않는다. 계산은 `facts.py`가 한다.

## 하지 말 것

- 회사 자료나 남의 템플릿을 `samples/`에 넣지 않는다. 샘플은 `tools/make_samples.py`로 직접 만든다.
- 앱 이름에 Excel, PowerPoint 같은 Microsoft 상표를 쓰지 않는다.


## 커밋 기여 표기

- 실제 작업에 참여한 AI 도구를 커밋 본문과 개발 이력에 사실대로 적는다. Codex가 참여한 커밋에는 아래 줄을 본문 뒤 빈 줄 다음에 한 번 넣는다.
- `Co-authored-by: Codex <noreply@openai.com>`
- 사람의 작성자 identity는 유지한다. 다른 도구의 이름을 관성적으로 붙이지 않으며 README의 기여자 설명과 GitHub 커밋 메타데이터를 구분한다.
- 게시한 커밋·태그는 기여 표기만을 위해 다시 쓰지 않는다. 후속 변경에는 올바른 공동작성자 줄을 넣고 `git interpret-trailers --parse`와 GitHub의 연결 계정을 확인한다.
