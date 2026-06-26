# ChatGPT Pro 의사결정 요청서: Local GPT for PowerPoint v0.3 이후

작성일: 2026-06-25  
프로젝트 경로: `C:\Users\saman\Documents\Codex\2026-06-25\gpt-for-powerpoint\outputs\local-gpt-powerpoint`

## 1. 현재까지 구현된 상태

Local GPT for PowerPoint는 Office 2016에서도 동작하도록 로컬 서버 + PowerPoint COM 자동화 방식으로 구현되어 있다.

현재 완료된 핵심 기능:

- ChatGPT/Codex OAuth 기반 사용
- OpenAI API key 입력 없음
- Codex 현재 모델 설정 사용
- PowerPoint context 읽기
- `SlideID` 포함 snapshot
- `deckFingerprint` 생성
- `PresentationPlan -> ExecutionPlan -> bridge` 구조
- strict JSON schema 기반 plan 검증
- legacy action allowlist
- read-only/review guard
- outline card UI
- 선택한 change만 preview/commit
- shadow deck preview
- before/after PNG preview
- transaction record 저장
- commit 전 `SaveCopyAs` backup
- 중복 commit idempotency
- rollback은 현재 backup copy를 새 PowerPoint로 여는 방식
- 업무 덱에는 preview만 수행했고, 실제 commit 테스트는 disposable test deck에서만 수행

## 2. Pro가 결정해줘야 하는 핵심 질문

아래 항목은 Codex가 계속 구현하기 전에 ChatGPT Pro가 방향을 정해주면 좋은 의사결정 목록이다.

## 3. Architecture 결정

### 결정 1: legacy action을 얼마나 오래 유지할지

현재는 모델이 `legacyActions`를 만들고, 서버가 이것을 `ExecutionPlan`으로 감싼 뒤 bridge에 전달한다.

결정 필요:

- 다음 단계에서도 `legacyActions`를 계속 호환 계층으로 둘 것인가?
- 아니면 `text.set`, `slide.insert_from_layout`, `notes.set` 같은 primitive operation으로 완전히 전환할 것인가?
- 전환한다면 어떤 순서가 안전한가?

Pro에게 물어볼 질문:

```text
현재 Local GPT for PowerPoint는 legacyActions를 ExecutionPlan으로 감싸는 과도기 구조입니다. Office 2016 COM bridge 안정성을 고려할 때, legacyActions를 얼마나 오래 유지하고 어떤 primitive operation부터 전환하는 것이 좋습니까?
```

## 4. Rollback 정책 결정

### 결정 2: rollback을 inverse operation 중심으로 갈지, backup 복구 중심으로 갈지

현재 rollback은 원본 파일을 자동 덮어쓰지 않고, backup copy를 새 프레젠테이션으로 여는 방식이다.

결정 필요:

- inverse operation rollback을 적극 구현할 것인가?
- 아니면 backup copy 기반 복구를 기본 정책으로 유지할 것인가?
- 사용자가 “되돌리기”를 누르면 원본을 자동 복구해도 되는가?
- 자동 복구를 한다면 원본 overwrite 허용 조건은 무엇인가?

Pro에게 물어볼 질문:

```text
PowerPoint COM 자동화 환경에서 rollback을 inverse operation으로 구현하는 것과 backup copy를 여는 복구 방식 중 어느 쪽을 기본 정책으로 삼는 것이 안전합니까? 원본 파일을 자동 덮어쓰는 복구는 허용해야 합니까?
```

## 5. Preview 정책 결정

### 결정 3: shadow deck preview를 얼마나 엄격하게 할지

현재 preview는 shadow deck에 적용하고 before/after PNG를 만든다. 다만 Office 2016 호환 때문에 shadow deck을 완전 hidden으로 열지 못하고 일반 `Open(path)` 후 닫는다.

결정 필요:

- 이 preview 방식을 v0.3 기본으로 인정할 것인가?
- hidden open 실패 가능성을 감수하고 계속 시도할 것인가?
- preview 중 PowerPoint 창 깜빡임이 있어도 괜찮은가?
- preview가 실패하면 commit을 무조건 막는 것이 맞는가?

Pro에게 물어볼 질문:

```text
Office 2016 COM에서 shadow deck을 완전 hidden으로 안정적으로 열기 어렵습니다. 일반 Open(path) 후 닫는 방식으로 preview를 구현해도 제품 UX상 허용 가능한가요? Preview 실패 시 commit을 무조건 막는 정책이 맞나요?
```

## 6. Conflict 검사 결정

### 결정 4: deck fingerprint 전체 기준 vs affected slide 기준

현재는 affected slide fingerprint를 우선 보지만, deck fingerprint도 기록한다.

결정 필요:

- 다른 슬라이드가 바뀌어도 대상 슬라이드만 그대로면 commit을 허용할 것인가?
- 아니면 plan 생성 후 deck 전체 fingerprint가 바뀌면 무조건 conflict로 볼 것인가?
- 업무용 PPT에서 협업/수동 수정 가능성을 어떻게 처리할 것인가?

Pro에게 물어볼 질문:

```text
Plan 생성 후 다른 슬라이드만 변경되고 대상 슬라이드는 그대로인 경우 commit을 허용해야 합니까? 아니면 deck 전체 fingerprint가 달라지면 항상 conflict로 막아야 합니까?
```

## 7. Selection 작업 결정

### 결정 5: 선택 영역 편집을 어떻게 preview할지

현재 `format_selection`은 live selection에 의존하므로 shadow preview에서 안전하게 재식별하기 어렵다.

결정 필요:

- 선택 영역 작업은 preview 없이 바로 backup 후 commit할 것인가?
- 선택한 shape에 `LOCALGPT_ID` tag를 먼저 붙이는 별도 승인 단계를 만들 것인가?
- selection fingerprint가 바뀌면 commit을 막는 정책으로 충분한가?

Pro에게 물어볼 질문:

```text
PowerPoint selection 기반 작업은 shadow deck에서 동일 shape를 재식별하기 어렵습니다. 선택 영역 편집은 preview 없이 backup 후 commit해도 되는지, 아니면 shape tagging/preflight 단계를 먼저 둬야 하는지 결정해 주세요.
```

## 8. High-Risk Operation 결정

### 결정 6: `replace_deck`을 언제 허용할지

`replace_deck`은 전체 덱 교체라 rollback 부담이 크다.

결정 필요:

- v0.3에서는 `replace_deck` live commit을 막을 것인가?
- preview와 backup이 있으면 허용할 것인가?
- 최소 어떤 warning/confirm UI가 필요할까?

Pro에게 물어볼 질문:

```text
전체 덱을 바꾸는 replace_deck operation은 v0.3에서 live commit을 막아야 합니까? 허용한다면 어떤 경고, preview, backup, confirmation 조건이 필요합니까?
```

## 9. UI/UX 결정

### 결정 7: 기본 모드를 `편집`으로 둘지 `검토만`으로 둘지

현재 UI는 `편집` 모드가 기본 선택이다. 하지만 안전성 관점에서는 `검토만`이 더 보수적이다.

결정 필요:

- 기본값을 `편집`으로 둘 것인가?
- `검토만`으로 두고 사용자가 명시적으로 편집을 선택하게 할 것인가?
- 공식 ChatGPT for PowerPoint와 비슷한 UX는 어느 쪽인가?

Pro에게 물어볼 질문:

```text
PowerPoint add-in의 기본 모드를 편집으로 둘지, 검토만으로 둘지 결정해 주세요. 안전성과 사용성, 공식 ChatGPT for PowerPoint의 기대 UX를 기준으로 판단해 주세요.
```

## 10. Backup 보존 정책 결정

### 결정 8: backup을 얼마나 보관할지

현재 backup은 `%LOCALAPPDATA%\LocalGptPowerPoint\backups`에 저장된다.

결정 필요:

- 최근 N개만 보관할 것인가?
- 30일 보관 정책을 둘 것인가?
- failed/recovery_required transaction backup은 영구 보존할 것인가?
- UI에서 backup 폴더 열기 기능을 넣을 것인가?

Pro에게 물어볼 질문:

```text
Local GPT PowerPoint backup 보존 정책을 정해 주세요. 최근 개수 기준, 기간 기준, failed transaction 예외 보존 기준을 어떻게 잡는 것이 좋습니까?
```

## 11. Native Chart/Data 결정

### 결정 9: 다음 버전에서 CSV/XLSX와 native chart 중 무엇부터 할지

v0.5 로드맵에는 CSV/XLSX deterministic parser와 PowerPoint native chart가 있다.

결정 필요:

- 먼저 CSV/XLSX parser를 만들 것인가?
- 먼저 native chart object 삽입을 만들 것인가?
- 데이터 계산은 모델이 아니라 deterministic parser가 해야 한다는 원칙을 어떻게 schema에 반영할 것인가?

Pro에게 물어볼 질문:

```text
다음 단계에서 CSV/XLSX parser와 PowerPoint native chart 중 무엇을 먼저 구현해야 합니까? 모델은 query specification만 제안하고 계산은 deterministic parser가 수행하는 구조를 어떻게 설계하면 좋습니까?
```

## 12. Template 보존 결정

### 결정 10: 기존 template/master를 어떻게 활용할지

현재 새 슬라이드는 기본 스타일로 생성된다.

결정 필요:

- 기존 slide를 duplicate해서 내용만 교체하는 방식으로 갈 것인가?
- CustomLayout 기반으로 새 슬라이드를 생성할 것인가?
- 회사 템플릿 보존을 위해 어떤 `TemplateProfile`을 먼저 수집해야 하는가?

Pro에게 물어볼 질문:

```text
기존 회사 PPT 템플릿을 보존하려면 slide duplicate 방식과 CustomLayout 방식 중 무엇을 우선해야 합니까? Office 2016 COM 환경에서 TemplateProfile은 어떤 정보부터 수집해야 합니까?
```

## 13. Connector 결정

### 결정 11: 외부 자료 연결을 taskpane OAuth로 할지 Codex connector로 할지

현재는 로컬 파일 업로드 중심이다.

결정 필요:

- Google Drive/OneDrive/Notion/Gmail을 taskpane 내부 OAuth로 붙일 것인가?
- Codex connector를 통해 server-side read-only source로 붙일 것인가?
- connector는 read-only부터 시작하는 것이 맞는가?

Pro에게 물어볼 질문:

```text
Google Drive, OneDrive, Notion, Gmail 같은 외부 자료 연결을 PowerPoint taskpane 내부 OAuth로 구현하는 것이 좋습니까, 아니면 Codex connector/server-side read-only source 방식이 좋습니까?
```

## 14. Test Strategy 결정

### 결정 12: 실제 업무 덱 테스트 허용 범위

현재 실제 업무 덱에는 preview만 수행했고 commit은 하지 않았다.

결정 필요:

- 실제 업무 덱에 commit 테스트를 언제 허용할 것인가?
- 허용한다면 어떤 조건이 필요한가?
- disposable test deck, copied real deck, live real deck의 단계 구분을 어떻게 할 것인가?

Pro에게 물어볼 질문:

```text
실제 업무 PPT에 live commit 테스트를 언제 허용해야 합니까? disposable deck, copied real deck, live real deck의 테스트 단계를 어떻게 나누는 것이 안전합니까?
```

## 15. Pro에게 그대로 붙여넣을 통합 프롬프트

```text
Local GPT for PowerPoint를 만들고 있습니다.

환경:
- Windows
- PowerPoint Office 2016 Professional Plus
- 공식 ChatGPT for PowerPoint add-in은 sideload는 됐지만 내부 web app이 로드되지 않음
- OpenAI API key 사용 불가
- Codex/ChatGPT OAuth만 사용
- Codex 현재 모델 설정 사용

현재 구현:
- localhost HTTPS taskpane
- PowerPoint COM bridge
- PresentationPlan -> ExecutionPlan -> bridge 구조
- strict JSON schema 검증
- SlideID snapshot
- deckFingerprint
- outline card UI
- selected change preview
- shadow deck preview
- before/after PNG preview
- transaction record
- commit 전 SaveCopyAs backup
- idempotent commit
- rollback은 backup copy를 새 PowerPoint로 여는 복구 방식

업무 덱에는 preview만 수행했고 live commit은 하지 않았습니다.
Disposable test deck에서는 commit, backup, idempotency, backup open rollback 경로를 확인했습니다.

이제 다음 의사결정이 필요합니다.

1. legacyActions를 계속 유지할지 primitive operation으로 완전히 전환할지
2. rollback을 inverse operation 중심으로 갈지 backup 복구 중심으로 갈지
3. Office 2016에서 shadow deck을 일반 Open(path) 후 닫는 preview 방식을 허용할지
4. conflict 검사를 deck 전체 기준으로 할지 affected slide 기준으로 할지
5. selection 기반 edit을 preview 없이 backup 후 commit해도 되는지
6. replace_deck 같은 high-risk operation을 v0.3에서 막을지
7. 기본 UI 모드를 편집으로 둘지 검토만으로 둘지
8. backup 보존 정책을 어떻게 정할지
9. CSV/XLSX parser와 native chart 중 무엇을 먼저 할지
10. template/master 보존은 duplicate exemplar와 CustomLayout 중 무엇부터 할지
11. connector는 taskpane OAuth와 Codex connector 중 어떤 구조가 나은지
12. 실제 업무 덱 live commit 테스트는 어떤 조건에서 허용할지

각 항목에 대해 권장 결정, 이유, 구현 우선순위, 위험 요소를 정리해 주세요.
```

## 16. Codex가 다음에 구현하기 전 필요한 우선 결정 5개

가장 먼저 결정하면 개발 속도가 빨라지는 항목은 다음 5개다.

1. 기본 모드: `편집` vs `검토만`
2. rollback 정책: inverse operation vs backup copy 복구
3. conflict 정책: affected slide 기준 vs deck 전체 기준
4. `replace_deck` live commit 허용 여부
5. selection edit preview 정책

