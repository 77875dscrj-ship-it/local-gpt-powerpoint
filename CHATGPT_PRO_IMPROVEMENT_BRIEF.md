# ChatGPT Pro 문의용 브리프: Local GPT for PowerPoint

작성일: 2026-06-25  
프로젝트 위치: `C:\Users\saman\Documents\Codex\2026-06-25\gpt-for-powerpoint\outputs\local-gpt-powerpoint`

## 1. 문의 목적

이 문서는 ChatGPT Pro에게 현재까지 만든 PowerPoint용 Local GPT Add-in의 구현 상태를 공유하고, 앞으로 공식 `ChatGPT for PowerPoint`에 가까운 기능으로 발전시키기 위한 개선 방향을 묻기 위한 자료입니다.

핵심 목표는 다음과 같습니다.

- 오래된 Office 2016 환경에서도 동작하는 PowerPoint용 AI 보조 도구를 로컬 방식으로 구현한다.
- 사용자는 OpenAI API 키를 넣을 수 없으므로, API key 방식이 아니라 ChatGPT/Codex OAuth 로그인 상태를 활용한다.
- GPT 모델은 별도 하드코딩하지 않고, 현재 Codex가 사용 중인 모델 옵션을 따라가게 한다.
- 단순 채팅 답변이 아니라 실제 PowerPoint 문서를 읽고, 계획을 만들고, 사용자가 허용하면 슬라이드를 직접 편집한다.

## 2. 배경과 제약

### 현재 PC/Office 환경

- Windows 로컬 PC에서 작업 중이다.
- PowerPoint는 Office 2016 Professional Plus 계열이며 확인된 버전은 `16.0.4266.1001`이다.
- 공식 OpenAI `ChatGPT for PowerPoint` Add-in은 sideload 자체는 되었지만, 내부 웹앱이 로드되지 않아 실제 사용 가능한 상태까지 가지 못했다.
- 원인은 오래된 Office 2016 Web Add-ins runtime/WebView 호환성 문제로 판단된다.
- PowerPoint 내부에서 `OfficeExtensionsDialog` 같은 웹 Add-ins 명령이 `Enabled=False`로 잡힌 정황도 있었다.

### 반드시 지켜야 하는 조건

- OpenAI API 키 입력 방식은 사용할 수 없다.
- ChatGPT/Codex OAuth 기반이어야 한다.
- GPT 모델은 현재 Codex 실행 옵션을 따라야 한다.
- PowerPoint에서 만든 결과물은 이미지가 아니라 가능한 한 편집 가능한 네이티브 텍스트 박스/도형/표 형태여야 한다.
- 오래된 Office 2016에서도 최대한 동작해야 하므로 최신 Office.js 기능에 의존하기 어렵다.

## 3. 현재 구현된 구조

현재 구현은 공식 Add-in을 억지로 고치는 방식이 아니라, 로컬 서버와 PowerPoint COM 자동화를 이용하는 우회 구조다.

### 주요 구성 파일

- `server.js`
  - HTTPS 로컬 서버
  - 정적 taskpane UI 제공
  - Codex CLI 호출
  - PowerPoint 브리지 PowerShell 스크립트 호출
  - `/api/health`, `/api/ppt/context`, `/api/source/extract`, `/api/chat/plan`, `/api/chat/apply` 제공

- `public/taskpane.html`
  - PowerPoint 작업창에 표시되는 Local GPT UI

- `public/taskpane.js`
  - 채팅 UI 로직
  - 파일 업로드
  - PowerPoint 컨텍스트 요청
  - AI 계획 생성 요청
  - 편집 허용 후 적용 요청

- `public/styles.css`
  - Office 2016 WebView에서도 동작하도록 단순한 CSS 위주

- `scripts/ppt-bridge.ps1`
  - PowerPoint COM 자동화 핵심
  - 열린 PowerPoint 인스턴스를 찾아 현재 프레젠테이션을 읽고 수정

- `scripts/extract-source.ps1`
  - 업로드 파일 텍스트 추출
  - TXT/MD/CSV/TSV/JSON/XML/HTML/LOG/PDF/DOCX/PPTX/XLSX 대응

- `scripts/ensure-cert.ps1`
  - `https://localhost:8765`용 로컬 인증서 생성/설치

- `manifest.xml`
  - Office Add-in sideload용 manifest

- `start-local-gpt.cmd`
  - 로컬 서버 실행용 Windows CMD 래퍼

## 4. 현재 동작 흐름

1. 사용자가 `start-local-gpt.cmd`로 로컬 서버를 실행한다.
2. PowerPoint에서 manifest 기반 Add-in 작업창을 연다.
3. 작업창은 `https://localhost:8765`의 Local GPT UI를 표시한다.
4. 사용자가 자연어로 요청을 입력한다.
5. 서버는 현재 PowerPoint 문맥을 읽는다.
   - 전체 슬라이드 수
   - 현재 슬라이드 번호
   - 현재 슬라이드 텍스트
   - 전체 덱 텍스트
   - 선택된 텍스트/도형 정보
6. 서버는 Codex CLI를 통해 ChatGPT OAuth 기반 모델에 계획 생성을 요청한다.
7. 모델은 `assistantMessage`, `steps`, `actions` 형태의 JSON 계획을 반환한다.
8. UI는 먼저 계획을 보여주고 `프레젠테이션 편집 허용` 버튼을 표시한다.
9. 사용자가 허용하면 `/api/chat/apply`가 호출된다.
10. `ppt-bridge.ps1`이 PowerPoint COM으로 실제 슬라이드를 수정한다.

## 5. 현재 구현된 기능

### 채팅/UX

- PowerPoint 작업창 안에서 채팅형 UI 제공
- 현재 상태 표시
- 새로고침 버튼
- 파일 추가 버튼
- AI가 편집 액션을 만들 경우 바로 적용하지 않고 사용자 허용을 요구
- 채팅 영역 높이 개선
  - 이전에는 채팅창이 너무 짧아 긴 대화를 보기 어려웠다.
  - 현재는 작업창 높이에 맞춰 채팅 영역이 자동 확장된다.
- Enter 키 전송 지원
  - `Enter`: 메시지 전송
  - `Shift+Enter`: 줄바꿈

### PowerPoint 읽기

- 열린 PowerPoint 앱 탐지
- 현재 프레젠테이션 확인
- 현재 슬라이드 번호/전체 슬라이드 수 확인
- 현재 슬라이드 텍스트 읽기
- 전체 덱 텍스트 읽기
- 선택된 도형/텍스트 정보 읽기
  - 도형 이름/ID
  - 위치/크기
  - 텍스트
  - 폰트 크기

### PowerPoint 편집

현재 지원되는 액션 타입은 다음과 같다.

- `replace_deck`
  - 전체 프레젠테이션을 새 슬라이드 구성으로 교체

- `add_slides`
  - 현재 슬라이드 뒤 또는 지정 위치 뒤에 새 슬라이드 추가

- `set_title`
  - 현재 또는 지정 슬라이드 제목 수정

- `set_body`
  - 현재 또는 지정 슬라이드 본문 수정

- `set_notes`
  - 발표자 노트 수정

- `replace_text`
  - 슬라이드 내 특정 텍스트 찾아 바꾸기

- `format_selection`
  - 선택된 텍스트/도형의 텍스트, 폰트 크기, 위치, 크기, 채우기 등을 수정

- `add_table_slide`
  - 편집 가능한 표 형태의 슬라이드 추가

- `add_bar_chart_slide`
  - 도형 기반 막대그래프 슬라이드 추가

### 슬라이드 생성

- 제목 슬라이드
- 섹션 슬라이드
- 일반 콘텐츠 슬라이드
- 비교형 슬라이드
- 표 슬라이드
- 단순 막대그래프 슬라이드
- 결과물은 가능한 한 PowerPoint 네이티브 텍스트 박스/도형 기반으로 생성

### 파일 기반 작업

작업창에서 파일을 추가하면 서버가 텍스트를 추출해 이후 프롬프트 문맥에 포함한다.

지원 형식:

- `.txt`
- `.md`
- `.csv`
- `.tsv`
- `.json`
- `.xml`
- `.html`
- `.htm`
- `.log`
- `.pdf`
- `.docx`
- `.pptx`
- `.xlsx`

## 6. 모델/OAuth 처리 방식

중요한 설계 결정:

- 서버는 OpenAI API key를 직접 사용하지 않는다.
- 서버는 `codex exec`를 호출한다.
- Codex CLI가 이미 보유한 ChatGPT OAuth 로그인 상태를 이용한다.
- 서버는 모델명을 별도로 강제하지 않는다.
- 즉, GPT 모델은 현재 Codex가 사용 중인 옵션을 따라간다.

최근 확인된 health 상태 예시는 다음과 같다.

```json
{
  "ok": true,
  "port": 8765,
  "authMode": "codex-chatgpt-oauth",
  "modelMode": "codex-current-option"
}
```

## 7. 검증된 항목

현재까지 확인한 테스트는 다음과 같다.

- `server.js` Node 문법 검사 통과
- `public/taskpane.js` Node 문법 검사 통과
- `scripts/ppt-bridge.ps1` PowerShell 파서 검사 통과
- `/api/health` 정상 응답 확인
- `/api/ppt/context`로 PowerPoint 컨텍스트 읽기 확인
- `/api/chat/plan`으로 AI 계획 생성 확인
- `/api/chat/apply`로 계획 적용 확인
- 선택된 텍스트 박스의 폰트 크기/위치/크기 수정 확인
- `replace_deck`으로 편집 가능한 새 슬라이드 생성 확인
- 파일 추가 후 텍스트 추출 확인
- 브라우저 기반 UI 테스트에서 다음 확인
  - UI 로드
  - 콘솔 오류 없음
  - 채팅 영역 자동 확장
  - Enter 전송
  - Shift+Enter 줄바꿈

## 8. 공식 ChatGPT for PowerPoint/영상에서 참고한 기능

참고한 자료:

- 공식 페이지: `https://chatgpt.com/ko-KR/apps/powerpoint/`
- OpenAI 영상: `https://www.youtube.com/watch?v=StEQpb3-S44`
- OpenAI 영상: `https://www.youtube.com/watch?v=ECkHTfvf2e8`
- The AI Advantage 영상: `https://www.youtube.com/watch?v=WJiHng4ymqw`

영상/공식 기능에서 관찰한 핵심 동작:

- PowerPoint 오른쪽 사이드바에서 ChatGPT와 대화
- 기존 프레젠테이션을 읽고 질의응답
- 소스 파일을 바탕으로 새 덱 생성
- 연결된 앱/파일을 참고해 콘텐츠 생성
- 생성 전 outline/plan을 먼저 보여줌
- 사용자가 허용하면 슬라이드를 직접 편집
- 결과물은 PowerPoint에서 직접 수정 가능한 텍스트/도형으로 생성
- 기존 덱 일부를 건드리지 않고 새 슬라이드만 추가
- 선택된 슬라이드 또는 선택된 텍스트 박스를 간결하게 수정
- 발표자 노트/구조/메시지 흐름 개선
- 약점, 누락 데이터, 예상 질문을 검토
- `Executive Readiness Quality Pass`처럼 목적별 품질 검토 수행
- CSV/분석 데이터에서 주요 수치를 뽑아 표/차트 슬라이드 생성

## 9. 현재 Add-in에서 이미 구현 가능하거나 거의 가능한 기능

다음은 현재 구조로 이미 가능하거나 약간의 UI/프롬프트 조정만으로 가능한 기능이다.

- 채팅 사이드바
- ChatGPT OAuth 기반 응답
- 현재 Codex 모델 옵션 사용
- 현재 PowerPoint 덱 읽기
- 현재 슬라이드/선택 영역 읽기
- 파일 업로드 기반 덱 생성
- 새 슬라이드 추가
- 전체 덱 재작성
- 제목/본문/노트 수정
- 선택된 텍스트 박스의 문구/크기/위치 조정
- 간단한 표 슬라이드 생성
- 도형 기반 막대그래프 생성
- 덱 리뷰/비판/개선점 제안
- 사용자가 허용하기 전에는 편집하지 않는 플로우

## 10. 아직 부족한 기능과 한계

### 공식 Add-in 수준으로 가려면 필요한 부분

- 슬라이드별 outline preview UI
  - 현재는 steps/actions 기반 계획은 있지만, 공식처럼 슬라이드별 미리보기 카드가 부족하다.

- 편집 diff/preview
  - 어떤 슬라이드가 어떻게 바뀌는지 더 명확히 보여줘야 한다.

- undo/rollback
  - 적용 전 자동 백업 또는 이전 상태 복원이 필요하다.

- 단일 슬라이드 polish 명령
  - “현재 슬라이드 텍스트가 너무 많으니 줄여줘” 같은 작업을 더 안정적으로 처리해야 한다.

- 선택 영역 기반 직접 수정
  - 현재 구현은 가능하지만 UX가 더 좋아져야 한다.

- 정교한 표/차트
  - 현재 막대그래프는 PowerPoint 도형 기반이다.
  - 진짜 PowerPoint Chart object 삽입은 아직 구현 전이다.

- 이미지/아이콘 삽입
  - 현재 이미지 생성, 웹 이미지 검색, 아이콘 삽입은 구현되지 않았다.

- 템플릿/브랜드 스타일 보존
  - 현재는 기본 스타일의 새 슬라이드 생성에 가깝다.
  - 기존 회사 템플릿의 master/theme/font/color를 정교하게 따르지는 못한다.

- 연결 앱
  - Google Drive, Microsoft 365, Gmail, Calendar, Notion 같은 외부 연결은 taskpane 내부에 아직 없다.
  - 현재는 로컬 파일 업로드 중심이다.

- 시각적 품질 평가
  - 현재는 텍스트/도형 메타데이터 중심이다.
  - 실제 화면 렌더링을 보고 레이아웃 겹침, 여백, 시각적 균형을 평가하는 기능은 부족하다.

### 기술적 한계

- Office 2016 Web Add-ins runtime이 오래되어 최신 Office.js/React 번들/ES module 의존이 위험하다.
- PowerPoint COM 자동화는 강력하지만 PowerPoint가 실행 중이어야 하며, 사용자 PC 상태에 영향을 받는다.
- Codex CLI를 OAuth 브리지로 쓰는 구조라 응답 속도와 세션 상태는 Codex 로그인 상태에 의존한다.
- 외부 SaaS connector를 붙이려면 별도 OAuth/권한/보안 설계가 필요하다.

## 11. 앞으로 업데이트할 기능 제안

### v0.2: UI/사용성 개선

- 슬라이드별 outline preview 카드 추가
- AI가 제안한 actions를 사람이 읽을 수 있는 변경 요약으로 표시
- 적용 전 “수정 대상 슬라이드”를 명확히 표시
- 채팅 히스토리 저장
- 자주 쓰는 작업 버튼 추가
  - 덱 만들기
  - 현재 슬라이드 줄이기
  - 임원 보고용으로 다듬기
  - 약점 찾기
  - 발표자 노트 만들기
- 적용 결과 로그 표시

### v0.3: 편집 액션 강화

- 현재 슬라이드만 다듬기
- 선택된 텍스트 박스만 다시 쓰기
- 선택된 도형 위치/크기 자동 정렬
- 특정 슬라이드 삭제
- 슬라이드 순서 변경
- 섹션 단위 추가/교체
- 제목 스타일 통일
- 본문 bullet 수 자동 축소
- 발표자 노트 일괄 생성/수정

### v0.4: 데이터/소스 기반 기능 강화

- CSV/XLSX를 구조화된 데이터로 파싱
- Top N 항목 자동 추출
- 수치 요약 자동 생성
- 데이터 기반 표 슬라이드 생성
- 데이터 기반 PowerPoint 차트 생성
- 출처 파일명/근거 문장을 발표자 노트에 남기기
- 여러 파일 간 충돌 내용 감지

### v0.5: 리뷰/품질검사 기능

- `Executive Readiness Quality Pass` 프리셋 구현
- 덱 약점 진단
- 누락 데이터 탐지
- 예상 질문 생성
- 슬라이드별 메시지 흐름 평가
- 제목만 읽어도 스토리가 이어지는지 검사
- 너무 긴 슬라이드/중복 bullet 탐지
- 숫자/단위/날짜 일관성 검사

### v0.6: 시각/템플릿 기능

- 기존 PPT theme/master/color/font 읽기
- 기존 템플릿과 어울리는 새 슬라이드 생성
- 실제 PowerPoint Chart object 삽입
- 이미지 삽입
- 아이콘 삽입
- 로컬 또는 웹 이미지 소스 연동
- 슬라이드 렌더링 기반 시각 QA

### v0.7: 안정성/배포

- 적용 전 자동 백업
- undo/rollback 버튼
- PowerPoint가 응답 없음 상태일 때 복구 안내
- 로컬 로그 파일
- 설치/실행 진단 화면
- Office 2016 전용 호환 모드
- Microsoft 365/Office 2021 이상 전용 고급 모드 분리
- 자동 업데이트 또는 재설치 스크립트

## 12. ChatGPT Pro에게 묻고 싶은 질문

아래 질문을 ChatGPT Pro에게 던져 개선 방향을 받으면 좋다.

1. Office 2016에서 최신 Office.js가 불안정한 상황이면, taskpane UI는 최소한으로 두고 PowerPoint COM 브리지를 중심으로 설계하는 것이 맞는가?
2. ChatGPT/Codex OAuth만 사용할 수 있고 API key는 사용할 수 없는 조건에서, `codex exec`를 브리지로 쓰는 현재 방식보다 더 안정적인 로컬 아키텍처가 있는가?
3. 공식 ChatGPT for PowerPoint와 비슷한 UX를 만들려면 action schema를 어떻게 재설계하는 것이 좋은가?
4. 슬라이드별 outline preview와 적용 전 diff preview는 어떤 데이터 구조로 만들면 좋은가?
5. PowerPoint COM으로 네이티브 차트, 표, 이미지, speaker notes, theme/master 스타일을 안정적으로 다루는 베스트 프랙티스는 무엇인가?
6. 적용 전 백업/undo/rollback을 구현할 때 가장 안전한 방식은 무엇인가?
7. 기존 회사 템플릿을 최대한 유지하면서 새 슬라이드를 추가하려면 어떤 정보를 읽고 어떤 규칙으로 레이아웃을 생성해야 하는가?
8. CSV/XLSX 같은 데이터 파일에서 주요 수치를 추출해 표/차트 슬라이드를 만들 때, 모델에게 맡길 부분과 deterministic parser로 처리할 부분을 어떻게 나누는 것이 좋은가?
9. deck review, quality pass, weak spot detection, missing data detection을 기능으로 만들 때 어떤 체크리스트와 출력 형식이 좋은가?
10. 연결 앱(Google Drive, Microsoft 365, Gmail, Notion 등)을 붙이려면 taskpane 내부 OAuth와 Codex connector 중 어떤 구조가 더 현실적인가?

## 13. ChatGPT Pro에 그대로 붙여넣을 프롬프트

아래 내용을 그대로 복사해서 ChatGPT Pro에게 물어보면 된다.

```text
나는 Windows 로컬 PC에서 PowerPoint용 Local GPT Add-in을 만들고 있다.

환경과 제약:
- PowerPoint는 Office 2016 Professional Plus 16.0.4266.1001이다.
- 공식 OpenAI ChatGPT for PowerPoint Add-in은 sideload는 됐지만 내부 웹앱이 로드되지 않았다.
- 원인은 오래된 Office 2016 Web Add-ins runtime/WebView 호환성으로 보인다.
- OpenAI API key를 넣을 수 없다.
- 반드시 ChatGPT/Codex OAuth 기반이어야 한다.
- GPT 모델은 현재 Codex가 사용하는 모델 옵션을 따라야 한다.
- 결과물은 이미지가 아니라 PowerPoint에서 직접 수정 가능한 텍스트 박스/도형/표/차트여야 한다.

현재 구현:
- https://localhost:8765 로컬 HTTPS 서버를 띄운다.
- PowerPoint Add-in manifest는 이 localhost taskpane을 연다.
- taskpane은 단순 HTML/CSS/JS 채팅 UI다.
- server.js가 Codex CLI의 codex exec를 호출해 ChatGPT OAuth 기반으로 응답을 받는다.
- server.js는 모델을 하드코딩하지 않고 현재 Codex 옵션을 따른다.
- PowerPoint 실제 수정은 scripts/ppt-bridge.ps1이 COM 자동화로 수행한다.
- 현재 PowerPoint context를 읽고, AI가 JSON plan/actions를 만들고, 사용자가 허용하면 적용한다.

현재 지원 actions:
- replace_deck
- add_slides
- set_title
- set_body
- set_notes
- replace_text
- format_selection
- add_table_slide
- add_bar_chart_slide

현재 가능:
- 현재 덱 읽기
- 현재 슬라이드/선택 영역 읽기
- 파일 업로드 후 텍스트 추출
- 전체 덱 생성/교체
- 새 슬라이드 추가
- 제목/본문/발표자 노트 수정
- 선택된 텍스트 박스 수정
- 표 슬라이드 생성
- 도형 기반 막대그래프 생성
- 계획 preview 후 사용자 허용을 받아 적용

앞으로 공식 ChatGPT for PowerPoint와 비슷하게 만들고 싶다.
특히 필요한 기능:
- 슬라이드별 outline preview
- 적용 전 diff preview
- undo/rollback
- 현재 슬라이드 polish
- 선택된 텍스트/도형 직접 수정
- CSV/XLSX 데이터 기반 표/차트 생성
- PowerPoint native chart object 삽입
- 기존 템플릿/theme/master 보존
- 이미지/아이콘 삽입
- Executive Readiness Quality Pass 같은 덱 품질검사
- 약점/누락 데이터/예상 질문 분석
- Google Drive/Microsoft 365/Gmail/Notion 같은 연결 앱 연동 가능성 검토

이 조건에서:
1. 전체 아키텍처를 어떻게 개선하는 것이 좋은가?
2. action schema를 어떻게 재설계해야 공식 add-in에 가까운 기능을 안정적으로 만들 수 있는가?
3. PowerPoint COM 자동화로 구현할 기능과 모델에게 맡길 기능을 어떻게 나눠야 하는가?
4. Office 2016 호환성을 유지하면서도 UX를 개선하려면 어떤 순서로 개발해야 하는가?
5. 우선순위별 개발 로드맵을 v0.2, v0.3, v0.4 식으로 제안해줘.
```

## 14. 우선순위 요약

가장 먼저 업데이트하면 효과가 큰 순서는 다음과 같다.

1. 슬라이드별 outline preview와 action summary UI
2. undo/backup/rollback
3. 현재 슬라이드/선택 영역 polish 기능
4. deck review/quality pass 프리셋
5. CSV/XLSX 기반 데이터 슬라이드
6. 기존 템플릿 스타일 보존
7. native chart/image/icon 삽입
8. 외부 connector 연동

