# ChatGPT Pro 검토 요청서: Local GPT for PowerPoint 하네스 통합 방향

작성일: 2026-06-26

## 1. 우리가 만들고 있는 것

우리는 **PowerPoint 안에서 AI가 직접 슬라이드를 읽고 수정할 수 있는 add-in**을 만들고 있다.

목표는 단순한 채팅창이 아니다. 사용자가 PowerPoint에서 슬라이드나 텍스트 박스, 도형 블록을 선택한 뒤 “이 블록 줄간격 넓혀줘”, “현재 슬라이드 내용 정리해줘”, “이 자료 기반으로 표 만들어줘”라고 지시하면, AI가 실제 PowerPoint 문서를 이해하고 직접 수정하는 add-in을 만드는 것이다.

다만 사용자의 환경은 **Office 2016 Professional Plus** 계열이다. 최신 Microsoft 365 PowerPoint처럼 최신 Office.js Runtime을 안정적으로 사용할 수 없다. 실제로 OpenAI의 공식 `ChatGPT for PowerPoint` add-in은 sideload 자체는 되었지만, 내부 웹앱이 PowerPoint 2016 taskpane에서 정상 로드되지 않았다.

그래서 현재 구현은 다음 구조를 사용한다.

1. PowerPoint 안에는 로컬 웹사이트 taskpane을 띄운다.
2. taskpane은 `https://localhost:8765`의 Local GPT 서버와 통신한다.
3. Local GPT 서버는 Codex/ChatGPT OAuth 기반 모델 호출을 사용한다. 사용자는 API 키를 넣지 않는다.
4. 실제 PowerPoint 수정은 PowerShell COM 자동화 브리지인 `scripts/ppt-bridge.ps1`이 수행한다.
5. AI는 바로 PowerPoint를 만지는 것이 아니라, 먼저 JSON 작업계획을 만들고, 서버가 그 계획을 검증한 뒤 PowerPoint에 적용한다.

즉 현재 구조는:

```text
PowerPoint 2016
  -> Local GPT taskpane HTML/JS
  -> Local Node server
  -> Codex/ChatGPT OAuth model call
  -> JSON edit plan
  -> PowerShell COM bridge
  -> PowerPoint native shapes/text/tables/charts
```

## 2. 현재 구현에 성공한 부분

현재까지는 실제로 PowerPoint 2016 안에 웹사이트 taskpane을 띄우는 데 성공했다.

구현된 주요 기능은 다음과 같다.

- PowerPoint 오른쪽 taskpane에 Local GPT UI 표시
- Codex/ChatGPT OAuth 기반 모델 호출
- 현재 프레젠테이션의 슬라이드 텍스트 읽기
- 현재 슬라이드의 도형, 텍스트 박스, 표, 차트 여부 일부 읽기
- 사용자가 선택한 텍스트 박스/도형 블록 인식
- 파일 추가 및 이미지 첨부
- PowerPoint taskpane에서 직접 이미지 붙여넣기가 안 되는 경우를 대비한 Windows 클립보드 이미지 읽기 fallback
- `replace_deck`, `add_slides`, `set_title`, `set_body`, `replace_text`, `format_selection`, `add_table_slide`, `add_bar_chart_slide` 같은 기본 작업계획 스키마
- PowerPoint COM을 통한 실제 편집 적용
- `Ctrl+Z`로 되돌릴 수 있도록 적용 전에 `StartNewUndoEntry()` 호출 시도

## 3. 현재 가장 큰 문제

현재 구현은 “붙이는 것”과 “수정하는 것” 자체는 가능하지만, **AI가 사용자의 지시를 정확하게 수행하는 능력이 부족하다.**

예를 들어 사용자가 선택한 본문 블록에 대해:

- “줄간격을 좀 넓혀줘”라고 했는데, 오히려 줄간격이 줄어드는 일이 있었다.
- “현재 선택한 블록을 수정해줘”라고 했는데, 선택 블록을 정확히 인식하지 못한 적이 있었다.
- “수정했다”고 응답했지만 실제 PowerPoint에서 눈에 띄는 변화가 없는 작업이 진행된 적이 있었다.
- 원본 보호 장치를 과하게 넣어둔 시기에는 미리보기와 복사본 작업이 반복되면서 실제 업무 흐름이 답답해졌다.

이 문제의 핵심은 모델 자체라기보다 **하네스(harness)의 부족**으로 보고 있다.

여기서 하네스란 AI가 PowerPoint를 제대로 조작할 수 있도록 잡아주는 장치 전체를 말한다.

예를 들면:

- 현재 선택한 도형이 정확히 무엇인지 알려주는 장치
- 현재 줄간격, 글꼴, 크기, 위치, 여백을 읽어주는 장치
- 슬라이드 마스터/레이아웃/테마를 읽어주는 장치
- “줄간격을 넓혀줘” 같은 상대적 지시를 현재 값 기준으로 계산해주는 장치
- AI가 만든 작업계획이 실제로 실행 가능한지 검증하는 장치
- 적용 후 정말 바뀌었는지 확인하는 장치
- 안 바뀐 작업은 성공으로 처리하지 못하게 하는 장치

현재 Local GPT는 PowerPoint를 “볼 수는 있지만, 충분히 잘 보지는 못한다.”  
그래서 조작하는 AI가 작업지시자의 의도를 명확히 수행하지 못하고 있다.

## 4. 왜 ChatGPT for PowerPoint / Claude for PowerPoint를 참고하려는가

우리는 공식 add-in 자체를 복사하려는 것이 아니다.  
목표는 **공식 add-in들이 PowerPoint를 이해하고 수정하기 위해 어떤 구조의 하네스를 쓰는지 참고하는 것**이다.

### 4.1 OpenAI ChatGPT for PowerPoint에서 확인한 것

OpenAI의 공식 `ChatGPT for PowerPoint` manifest는 확보했다.

확인된 구조:

- PowerPoint taskpane URL은 OpenAI 서버의 `bps.openai.com/.../powerpoint/`
- Office.js를 로드한다.
- `ReadWriteDocument` 권한을 사용한다.
- Word, Excel, PowerPoint를 모두 대상으로 하는 manifest 구조다.
- 내부에는 PowerPoint 전용 shell과 OpenAI 원격 gateway/service가 연결되어 있다.

다만 Office 2016에서는 taskpane shell은 붙었지만 내부 웹앱이 제대로 로드되지 않았다.

### 4.2 Anthropic Claude for PowerPoint에서 확인한 것

Anthropic의 `Claude for PowerPoint`도 조사했다.

확인된 구조:

- 공식 manifest URL: `https://pivot.claude.ai/manifest-powerpoint.xml`
- Taskpane URL: `https://pivot.claude.ai`
- Provider: `Anthropic`
- Host: `Presentation`
- Permission: `ReadWriteDocument`
- Requirement: `SharedRuntime 1.1`
- Shortcut override: `https://pivot.claude.ai/shortcuts.json`
- `Ctrl+Alt+C`로 sidebar를 열 수 있는 구조

Claude taskpane shell과 공개 JS bundle을 표면적으로 조사한 결과, 다음 단어들이 많이 등장했다.

- `slide`
- `shape`
- `layout`
- `chart`
- `deck`
- `master`
- `Office.context`
- `mcp`
- `skill`
- `gateway`
- `gateway_api_format`
- PowerPoint / Word / Excel 관련 OOXML namespace

이는 Claude add-in이 단순히 선택 텍스트만 보는 것이 아니라, **슬라이드, 도형, 레이아웃, 마스터, 테마, 차트, OOXML 구조까지 포함하는 더 풍부한 하네스**를 가지고 있을 가능성이 높다는 의미다.

## 5. 지금 고민하는 핵심: 기존 하네스와 Claude식 하네스가 꼬일 수 있음

중요한 우려가 있다.

우리는 이미 기존 하네스를 만들었다.

- PowerPoint COM bridge
- deck fingerprint
- shape fingerprint
- frozen selection
- JSON plan schema
- legacyActions
- preview / transaction / apply 흐름
- 선택 영역 기반 `format_selection`

여기에 Claude add-in에서 관찰한 하네스 단서를 그대로 덧붙이면 서로 꼬일 수 있다.

예를 들면:

1. 기존 `deckFingerprint`에 slide master/theme 정보를 갑자기 넣으면, 예전 계획들이 모두 충돌로 판단될 수 있다.
2. 기존 `legacyActions`에 Claude식 `mcp`, `skill`, `gateway` 개념을 섞으면, 모델이 실행 불가능한 작업을 만들 수 있다.
3. Office 2016에서는 Claude가 요구하는 `SharedRuntime 1.1` 방식이 안정적으로 작동하지 않을 수 있다.
4. 공식 add-in의 Office.js 기반 조작 방식과 우리 COM 기반 조작 방식을 섞으면 디버깅이 매우 어려워질 수 있다.
5. Claude/ChatGPT의 내부 원격 하네스 코드는 공개적으로 복사해서 사용할 수 있는 것이 아니며, 사용자의 OAuth 요구사항도 다르다.

따라서 현재 판단은 다음과 같다.

**Claude/OpenAI 하네스를 그대로 가져오는 것이 아니라, 관찰 계층만 참고해서 우리 하네스를 재설계해야 한다.**

## 6. 현재 생각하는 통합 원칙

### 원칙 1: 기존 COM bridge는 실제 조작 엔진으로 유지

Office 2016 환경에서는 PowerPoint COM 자동화가 가장 현실적인 조작 방식이다.

따라서 실제 수정 경로는 계속 다음처럼 유지한다.

```text
AI plan -> Local server validation -> PowerShell COM bridge -> PowerPoint edit
```

Claude/OpenAI식 Office.js 런타임은 그대로 섞지 않는다.

### 원칙 2: Claude/OpenAI 하네스는 “보는 눈”으로 먼저 적용

공식 add-in에서 참고할 것은 실행 코드가 아니라 “PowerPoint를 어떤 정보 단위로 이해하는가”이다.

우선 다음 정보를 읽기 전용 context로 추가한다.

- slide master 정보
- custom layout 정보
- design/theme 정보
- active slide의 layout name/index
- shape별 현재 font name/font size/bold/color
- paragraph lineSpacing/spaceBefore/spaceAfter
- shape별 table/chart/picture 여부
- OOXML 기반 추가 구조 정보

이 정보는 모델의 판단에만 사용하고, 바로 실행 스키마와 섞지 않는다.

### 원칙 3: fingerprint 계층을 분리

기존 `deckFingerprint`는 현재 변경 충돌 검사용으로 유지한다.

새로 읽는 master/layout/theme 정보는 별도 fingerprint로 둔다.

예:

```text
deckFingerprint       = 기존 슬라이드/텍스트/도형 변경 감지용
selectionFingerprint  = 선택 도형 변경 감지용
shapeFingerprint      = 개별 도형 재식별용
templateFingerprint   = 마스터/레이아웃/테마 변경 감지용
```

이렇게 해야 기존 transaction 흐름이 갑자기 깨지지 않는다.

### 원칙 4: action schema는 바로 확장하지 않음

Claude bundle에서 `mcp`, `skill`, `gateway`, `chart`, `layout` 흔적이 보였다고 해서 바로 우리 action schema에 넣으면 안 된다.

먼저 읽기 전용 context를 안정화한 뒤, 실제 작업 단위는 작게 추가한다.

후보:

```text
text.format
paragraph.format
shape.resize
shape.move
shape.fill
slide.insert_from_layout
table.update
chart.create_native
notes.set
```

### 원칙 5: 적용 후 검증을 반드시 강화

현재는 “작업계획이 적용됨”과 “실제로 눈에 띄는 변경이 있었음”이 완전히 분리되어 있지 않다.

앞으로는 다음 검증이 필요하다.

- 적용 전 snapshot
- 적용 후 snapshot
- 변경된 속성 목록
- 실제 변경 수
- no-op이면 성공 처리 금지
- 사용자의 원래 요청과 적용 결과가 맞는지 검토 메시지 표시

## 7. 이미 일부 시작된 변경

현재 `scripts/ppt-bridge.ps1`에 읽기 전용 context 확장 작업이 일부 들어갔다.

추가된 방향:

- `Get-TextStyleInfo`
- shape별 fontName/fontSize/bold/fontRgb
- paragraph lineSpacing/spaceBefore/spaceAfter
- slide layout/design name
- design/master/custom layout 정보
- template context

하지만 아직 Node 서버와 planner prompt에 완전히 연결하지는 않았다.

즉 현재 단계는 “실제 조작 엔진 변경”이 아니라, “PowerPoint를 더 잘 읽기 위한 정보 수집 확장”에 가깝다.

## 8. ChatGPT Pro에게 묻고 싶은 핵심 질문

### 질문 1: 이 통합 방향이 맞는가?

Office 2016 제약을 고려할 때, 공식 ChatGPT/Claude add-in의 Office.js 기반 실행 구조를 따라가기보다는, 기존 COM bridge를 유지하고 공식 add-in의 관찰/문맥 하네스만 참고하는 방향이 맞는가?

### 질문 2: 하네스 계층을 어떻게 나누는 것이 좋은가?

현재 생각은 다음 4계층이다.

```text
1. Observation Harness
   PowerPoint의 현재 상태를 읽음

2. Planning Harness
   AI가 사용자 지시를 실행 가능한 계획으로 바꿈

3. Execution Harness
   검증된 계획만 PowerPoint COM으로 적용

4. Verification Harness
   적용 후 실제 변경 여부를 확인
```

이 계층 구분이 적절한가? 더 좋은 구조가 있는가?

### 질문 3: 기존 deckFingerprint를 유지하고 templateFingerprint를 분리하는 것이 맞는가?

마스터/테마/레이아웃 정보까지 기존 `deckFingerprint`에 넣으면 충돌 감지가 너무 민감해질 수 있다.

따라서 `templateFingerprint`를 별도 필드로 두고, 처음에는 planning context에만 사용하려 한다.

이 접근이 맞는가?

### 질문 4: Claude/OpenAI add-in에서 참고해야 할 핵심 설계 포인트는 무엇인가?

우리가 공개적으로 관찰한 것은 manifest, taskpane shell, 공개 JS bundle의 표면 문자열, 공식 문서 수준이다.

이 정보에서 합법적이고 실용적으로 참고할 수 있는 설계 포인트는 무엇인가?

예:

- SharedRuntime 구조
- shortcut handling
- taskpane shell
- gateway/bootstrap config
- MCP/Skills 개념
- slide/master/layout/theme/OOXML-aware context

무엇을 우선 적용해야 하는가?

### 질문 5: `legacyActions`를 계속 유지해도 되는가?

현재 action schema는 `legacyActions` 중심이다.

하지만 공식 add-in급 조작을 하려면 더 작은 primitive operation으로 바꿔야 할 것 같다.

예:

```text
replace_text
format_selection
add_table_slide
add_bar_chart_slide
```

에서:

```text
shape.text.replace
shape.text.format
shape.paragraph.format
shape.bounds.set
slide.insert_from_layout
table.cells.update
chart.native.create
```

같은 구조로 가는 것이 맞는가?

### 질문 6: 줄간격 같은 상대 지시를 어떻게 처리해야 하는가?

현재 문제 중 하나는 사용자가 “줄간격을 넓혀줘”라고 했을 때 모델이 절대값 `1.15`, `1.25` 등을 임의로 선택하면서, 기존 값보다 작아지는 경우가 생길 수 있다는 점이다.

하네스가 현재 값을 읽고 다음처럼 처리해야 하는가?

```text
current lineSpacing = 1.40
user asks "넓혀줘"
target lineSpacing = current + 0.10 or current * 1.10
```

이 계산은 모델에게 맡겨야 하는가, 아니면 deterministic server-side helper가 해야 하는가?

### 질문 7: no-op 방지를 어떻게 설계해야 하는가?

현재 “적용 완료”라고 했지만 실제 변경이 없거나 눈에 보이지 않는 경우가 있었다.

앞으로는 action별로 변경 전후 값을 비교해야 할 것 같다.

예:

```json
{
  "changed": true,
  "changedProperties": ["lineSpacing", "spaceAfter"],
  "before": { "lineSpacing": 1.0, "spaceAfter": 0.0 },
  "after": { "lineSpacing": 1.2, "spaceAfter": 0.2 }
}
```

이런 방식이 충분한가? 시각적 변화까지 확인해야 하는가?

### 질문 8: Office 2016에서 taskpane clipboard/image/OCR/context를 어떻게 다루는 것이 좋은가?

PowerPoint 2016 taskpane은 오래된 WebView라 브라우저 Clipboard API가 제한될 수 있다.

현재는 Windows PowerShell로 클립보드 이미지를 읽는 fallback을 넣었다.

이 방식이 장기적으로 괜찮은가? 아니면 사용자가 선택한 슬라이드/도형을 이미지로 export해서 AI에게 함께 보내는 방식이 더 나은가?

### 질문 9: Claude의 MCP/Skills 개념을 Local GPT에 가져오는 것이 유용한가?

Claude bundle에는 `mcp`, `skill`, `gateway_api_format` 같은 흔적이 있다.

우리도 Local GPT 안에 PowerPoint용 local skills를 둘 수 있을 것 같다.

예:

```text
skill: slide_cleanup
skill: dense_text_relax
skill: table_from_source
skill: executive_summary_slide
skill: native_chart_from_csv
```

이런 skill layer를 planner prompt에 넣는 것이 좋은가? 아니면 먼저 primitive operation 안정화가 우선인가?

### 질문 10: 지금 당장 다음 작업 순서는 무엇이 좋은가?

현재 생각하는 순서는 다음이다.

1. 기존 하네스를 깨지 않도록 통합 규칙 문서 작성
2. `templateContext`를 읽기 전용으로 추가
3. `paragraph/style context`를 selection/shape map에 추가
4. planner prompt가 현재값 기준으로 판단하도록 수정
5. `paragraph.format` primitive operation 추가
6. no-op detector 강화
7. 적용 후 before/after property report 표시
8. 이후 layout/theme/master 기반 slide insertion 검토

이 순서가 맞는가?  
아니면 먼저 action schema를 갈아엎는 것이 맞는가?

## 9. ChatGPT Pro에게 원하는 답변 형식

다음 형식으로 답변해주면 좋겠다.

1. 현재 방향에 대한 총평
2. 기존 하네스와 Claude/OpenAI식 하네스를 섞을 때 가장 위험한 충돌 지점
3. 반드시 유지해야 할 기존 구조
4. 새로 추가해도 안전한 관찰 context
5. 지금 추가하면 안 되는 기능
6. 추천하는 하네스 아키텍처
7. 권장 action schema 예시
8. Office 2016 기준 현실적인 구현 순서
9. 줄간격/선택 블록/no-op 문제에 대한 구체적 해결책
10. 최종 권장 로드맵

## 10. 한 줄 요약

우리는 Office 2016에서도 동작하는 PowerPoint용 Local GPT add-in을 만들고 있으며, taskpane과 실제 수정은 구현했지만 AI가 PowerPoint를 충분히 잘 이해하지 못해 작업 지시를 정확히 수행하지 못하고 있다. 이를 해결하기 위해 공식 ChatGPT for PowerPoint와 Claude for PowerPoint의 하네스 구조를 참고하되, 기존 COM 기반 하네스와 충돌하지 않도록 “관찰 계층부터 분리해서 통합하는 방식”이 맞는지 ChatGPT Pro에게 검토받고 싶다.
