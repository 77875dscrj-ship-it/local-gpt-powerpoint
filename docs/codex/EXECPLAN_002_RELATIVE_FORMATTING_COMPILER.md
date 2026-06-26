# ExecPlan 002 - Deterministic relative selection formatting compiler

This ExecPlan follows `docs/codex/PLANS.md`. Keep the `Progress`, `Surprises & Discoveries`, `Decision Log`, and `Outcomes & Retrospective` sections updated as work proceeds.

## Purpose / Big Picture

ExecPlan 001 made selection formatting safer. It stopped the planner from guessing absolute formatting values such as `lineSpacing: 1.15` or `1.25`, preserved richer PowerPoint observation data, froze the selected shape at plan creation time, and reported no-op formatting as `변경 없음`.

That safety created a user-facing gap: a request like `선택한 블록의 줄간격을 조금 넓혀줘` is now refused unless the user gives an exact value. ExecPlan 002 should make this common request work safely by letting the server compute the exact target value from the current PowerPoint observation. A deterministic compiler means normal code, not the model, calculates the final number using fixed rules.

After this change, when a selected text shape has current `paragraph.lineSpacing = 1.00` and the user asks to slightly increase line spacing, the server should compile a `format_selection` action with a line spacing greater than `1.00`. When current `lineSpacing = 1.40`, the server must not choose a lower canned value like `1.15`; it should choose a deterministic greater value such as `1.50`.

This is not a primitive schema rewrite. Keep `legacyActions` as the adapter and keep the PowerShell COM bridge as the only PowerPoint write path.

## Progress

- [x] (2026-06-26 16:40 local) Initial investigation of ExecPlan 001 state, planner prompt, selection observation, and `format_selection` compile path complete.
- [x] (2026-06-26 16:40 local) ExecPlan 002 created with proposed helpers, file changes, deterministic rules, and test cases.
- [x] (2026-06-26 17:45 local) Implementation step complete in `server.js` only.
- [x] (2026-06-26 18:10 local) Validation complete with bundled Node checks, bridge parse check, diff check, server helper smoke, and disposable PowerPoint smoke.

## Surprises & Discoveries

- The working branch `execplan-002-relative-formatting` initially pointed at the initial commit and did not contain `AGENTS.md` or `docs/codex`. It was clean, so it was fast-forwarded to `codex-handoff-001` before this plan was written.
- ExecPlan 001 already normalized and preserved the fields needed for this task: `fontSize`, `paragraph.lineSpacing`, `paragraph.spaceAfter`, `hasTextFrame`, `hasTable`, `hasChart`, and frozen selection shape metadata.
- The existing presentation plan schema already allows `fontSize`, `lineSpacing`, and `spaceAfter` on `format_selection`, so this plan should not need a schema change.
- The current prompt tells the model not to invent absolute values for relative requests. ExecPlan 002 should keep that rule and add server-side compilation after the model response is normalized.
- Korean parsing needed a field-first rule: `줄간격` and `줄 간격` must identify the field, while only clear expressions like `줄여줘`, `작게`, or `좁혀` identify decrease direction.
- A normal PowerPoint text box generally returns a current line spacing value through COM, so the missing-current-value case was verified with a direct server-helper smoke rather than a live UI selection.
- The disposable PowerPoint smoke script was kept outside the repository runtime tree at `C:\Users\saman\Documents\Codex\2026-06-26\read-agents-md-docs-codex-plans\work\execplan002-relative-formatting-smoke.ps1`.

## Decision Log

- Decision: Add a small server-side compiler in `server.js`, not a new Office.js or PowerShell write path.
  Rationale: The current COM bridge already applies `format_selection` safely. The missing piece is deterministic target-value calculation before the existing bridge apply path.
  Date/Author: 2026-06-26 / Codex

- Decision: Keep `legacyActions` as the execution adapter.
  Rationale: `format_selection` already supports `fontSize`, `lineSpacing`, and `spaceAfter`; a primitive action schema rewrite is outside this task.
  Date/Author: 2026-06-26 / Codex

- Decision: Compute relative target values only from observed current values.
  Rationale: If the current value is missing, the server cannot safely calculate a relative target and must reject or require review instead of guessing.
  Date/Author: 2026-06-26 / Codex

- Decision: Use fixed small-step rules for the initial scope.
  Rationale: Deterministic, easy-to-test values are safer than model-selected or style-dependent guesses. Later tasks can add user-configurable or theme-aware steps.
  Date/Author: 2026-06-26 / Codex

- Decision: Do not change `deckFingerprint` semantics.
  Rationale: ExecPlan 001 kept template data separate as `templateFingerprint`; ExecPlan 002 should continue using `deckFingerprint` only for deck/content conflict checks.
  Date/Author: 2026-06-26 / Codex

- Decision: Detect relative formatting fields before directions, and never treat the bare character `줄` as a decrease direction.
  Rationale: In Korean, `줄` is part of the field name in `줄간격`; treating it as a decrease verb would invert requests like `줄간격을 조금 넓혀줘`.
  Date/Author: 2026-06-26 / Codex

- Decision: Compile line spacing only when the observed value is a sane multiple or explicitly has `lineRuleWithin = -1`.
  Rationale: PowerPoint can represent line spacing in point-based or unsupported modes; the compiler must not clamp unsafe values or guess a multiple.
  Date/Author: 2026-06-26 / Codex

- Decision: Let the server create the `format_selection` action in edit mode even when the model returned only review text.
  Rationale: ExecPlan 001 intentionally made the planner conservative. ExecPlan 002 closes the UX gap by computing exact values from observation in deterministic code.
  Date/Author: 2026-06-26 / Codex

- Decision: Extend the direct apply path to pass `message` and `requestedMode` into the same compiler.
  Rationale: This keeps `/api/chat/apply` smoke tests and direct legacy-plan calls aligned with the normal plan path without adding a new write path.
  Date/Author: 2026-06-26 / Codex

## Outcomes & Retrospective

Implemented ExecPlan 002 in `server.js` without changing `public/taskpane.js`, schemas, manifest, Codex OAuth, or `scripts/ppt-bridge.ps1`.

The compiler now detects supported relative selected-text formatting requests, validates that exactly one plain text shape is selected, reads the current value from `context.selection` with active-slide map fallback, computes deterministic target values, and emits a `format_selection` action through existing `legacyActions`.

Confirmed deterministic targets:

- `lineSpacing 1.00 -> 1.10` for `선택한 블록의 줄간격을 조금 넓혀줘`
- `lineSpacing 1.40 -> 1.50` for the same increase request
- `lineSpacing 1.40 -> 1.30` for `선택한 블록의 줄간격을 조금 줄여줘`
- `fontSize 18 -> 19` for `선택한 텍스트를 조금 크게 해줘`
- `fontSize 18 -> 17` for `선택한 텍스트를 조금 작게 해줘`
- `spaceAfter 6 -> 9` for a supported paragraph-after spacing increase

The safest limitation remains intentional: if the current value is missing, point-based, outside the sane line-spacing range, or the selection is table/chart/group/multiple/no shape, the compiler returns no edit action with a Korean explanation.

## Context and Orientation

The add-in has three relevant layers:

1. `public/taskpane.js` is the PowerPoint taskpane UI. It sends the user's message and edit/review mode to the server.
2. `server.js` reads PowerPoint context through `scripts/ppt-bridge.ps1`, builds the planner prompt, receives and validates a JSON plan, compiles `legacyActions` into an execution plan, and applies approved edits.
3. `scripts/ppt-bridge.ps1` is the only write path. It applies `format_selection` by setting native PowerPoint COM properties.

Relevant current `server.js` functions:

- `normalizeShape` preserves `fontSize`, `paragraph.lineSpacing`, `paragraph.spaceAfter`, table/chart flags, and other read-only fields.
- `augmentSelection` normalizes the current selection using `normalizeShape`.
- `shapeMapForPrompt` sends compact shape style data to the model.
- `buildPlannerPrompt` currently tells the model not to invent absolute values for relative line spacing or paragraph spacing.
- `normalizePresentationPlan` converts model output into the strict local plan shape.
- `validatePresentationPlan` checks the model plan.
- `compileExecutionPlan` freezes `format_selection` targets from the plan-time context.
- `annotateNoOpResult` marks zero-change formatting applies as no-op.

The deterministic compiler should live in `server.js` near the selection formatting helpers, before `compileExecutionPlan`. It should be called from `createPlan` after `normalizePresentationPlan` and before `validatePresentationPlan`, so the final plan that gets stored and shown to the taskpane contains server-computed exact values.

## Plan of Work

### 1. Add relative intent parsing in `server.js`

Add helper:

```js
function parseRelativeSelectionFormattingIntent(message) { ... }
```

It should return either `null` or an object like:

```js
{
  field: "lineSpacing" | "fontSize" | "spaceAfter",
  direction: "increase" | "decrease",
  strength: "small",
  source: "user_message"
}
```

Initial recognized language:

- Selection hints: `선택`, `블록`, `텍스트`, `selected`, `selection`.
- Font size field: `글자`, `글꼴`, `폰트`, `크기`, `font size`.
- Line spacing field: `줄간격`, `행간`, `line spacing`.
- Paragraph after spacing field: `문단 뒤`, `단락 뒤`, `문단 간격`, `paragraph spacing`, `space after`.
- Increase direction: `넓`, `늘`, `키`, `크게`, `increase`, `larger`, `more`.
- Decrease direction: `좁`, `줄`, `작게`, `decrease`, `smaller`, `less`.

If the message matches multiple fields or no direction, return an unsupported/ambiguous result that causes review or warning rather than a guessed edit.

### 2. Add selection validation in `server.js`

Add helper:

```js
function resolveRelativeFormattingSelection(context) { ... }
```

Rules:

- Require exactly one selected shape in `context.selection.shapes`.
- Require `hasTextFrame === true`.
- Reject `hasTable === true`.
- Reject `hasChart === true`.
- Reject grouped shapes. If only numeric `type` is available, treat PowerPoint/Office group type `6` as unsupported and document the limitation.
- Prefer current values from `context.selection.shapes[0]`.
- If a field is missing on the selection shape, optionally look up the same shape in `context.activeSlideShapeMap` by `shapeFingerprint`, `id`, or `name`.
- If the current value is still missing or not finite, reject instead of guessing.

### 3. Add deterministic value calculation in `server.js`

Add helpers:

```js
function currentRelativeFormatValue(shape, field) { ... }
function computeRelativeFormatTarget(field, currentValue, direction, strength) { ... }
function roundTo(value, decimals) { ... }
function clamp(value, min, max) { ... }
```

Initial deterministic rules:

- `fontSize`
  - Current unit: points.
  - Small step: `1.0`.
  - Bounds: `6` to `96`.
  - Round to `0.5` point or one decimal.
  - Increase target must be greater than current. Decrease target must be less than current.

- `lineSpacing`
  - Current unit: PowerPoint `ParagraphFormat.SpaceWithin`, as observed by the bridge.
  - Small step: `0.10`.
  - Bounds: `0.80` to `3.00`.
  - Round to two decimals.
  - Example: `1.00 -> 1.10` for increase.
  - Example: `1.40 -> 1.50` for increase.
  - Never choose a lower canned value such as `1.15` when current is already `1.40`.

- `spaceAfter`
  - Current unit: points.
  - Small step: `3.0`.
  - Bounds: `0` to `72`.
  - Round to one decimal.
  - Only compile when `paragraph.spaceAfter` is available.

If clamping would make the target equal to the current value, do not pretend to make a visible change. Return a no-op/review result with a clear Korean explanation such as `변경 없음: 이미 허용 범위의 끝 값입니다.`

### 4. Add server compiler integration in `server.js`

Add helper:

```js
function compileRelativeSelectionFormattingPlan(plan, options) { ... }
```

Suggested options:

```js
{
  message,
  context,
  requestedMode,
  readOnlyIntent
}
```

Behavior:

- If no relative formatting intent is detected, return the plan unchanged.
- If `requestedMode !== "edit"` or `readOnlyIntent === true`, keep review behavior and add a warning rather than an edit action.
- If the selection is unsupported or current value is missing, return a review/no-action plan with a clear Korean explanation.
- If the request is supported, set or replace the relevant `format_selection` action with server-computed absolute target values.
- Clear unrelated absolute formatting fields for this compiler-created action.
- Preserve `legacyActions` as the adapter.
- Add a matching `outline` item with the same `changeId`.
- Add a warning or detail noting the calculation, for example: `서버가 현재 줄간격 1.40에서 1.50으로 계산했습니다.`
- The model may describe relative intent, but the model must not choose the final absolute value. If the model returned a conflicting `format_selection` value for the same relative field, overwrite it with the server-computed value or replace the action.

Example compiler-created action:

```js
{
  type: "format_selection",
  changeId: "chg-relative-line-spacing",
  slide: null,
  after: null,
  text: null,
  find: null,
  replace: null,
  fontSize: null,
  width: null,
  height: null,
  left: null,
  top: null,
  autofit: null,
  lineSpacing: 1.5,
  spaceBefore: null,
  spaceAfter: null,
  bold: null,
  fillRgb: null,
  title: null,
  message: null,
  notes: null,
  columns: [],
  rows: [],
  items: [],
  slides: []
}
```

### 5. Update planner prompt in `server.js`

Keep the ExecPlan 001 safety rule, but add one sentence explaining the new server behavior:

- The model should express relative selected-block formatting intent without inventing numeric values.
- For supported selected text shape requests, the server will compute exact `fontSize`, `lineSpacing`, or `spaceAfter` values from current observation.

Do not ask the model to output `1.15`, `1.25`, or any other guessed absolute value for a relative request.

### 6. Keep taskpane changes minimal

Preferred initial plan: no required `public/taskpane.js` changes unless testing shows the no-action/review message is confusing. Existing no-op reporting from ExecPlan 001 should remain.

### 7. Do not edit these unless a test proves it is necessary

- `scripts/ppt-bridge.ps1`: existing `format_selection` already applies `fontSize`, `lineSpacing`, and `spaceAfter`.
- `schemas/presentation-plan.schema.json`: existing `format_selection` fields already cover this scope.
- `schemas/execution-plan.schema.json`: existing `legacyActions` adapter remains.
- `manifest.xml`: no SharedRuntime or Office.js write-path changes.
- Codex OAuth code: do not change authentication or model invocation flow.

## Concrete Steps

1. Confirm branch and status:

    ```powershell
    git status --short
    git branch --show-current
    ```

2. Implement only `server.js` helpers and integration:

    - Add `parseRelativeSelectionFormattingIntent`.
    - Add `resolveRelativeFormattingSelection`.
    - Add `currentRelativeFormatValue`.
    - Add `computeRelativeFormatTarget`.
    - Add `compileRelativeSelectionFormattingPlan`.
    - Call `compileRelativeSelectionFormattingPlan` from `createPlan` after `normalizePresentationPlan(raw, context)` and before `validatePresentationPlan(plan)`.
    - Add prompt wording that the server, not the model, computes exact values.

3. Leave `public/taskpane.js` unchanged unless validation proves a display issue.

4. Do not modify product code outside `server.js` unless the smallest possible fix requires it.

5. Update this ExecPlan as work proceeds.

6. Update `docs/codex/WORKLOG.md` after implementation and validation.

## Test Cases

### Parser and compiler behavior

1. Korean line spacing increase:
   - Message: `선택한 블록의 줄간격을 조금 넓혀줘`
   - Current: `paragraph.lineSpacing = 1.00`
   - Expected action: `format_selection.lineSpacing = 1.10`

2. Korean line spacing increase from higher current value:
   - Message: `선택한 블록의 줄간격을 조금 넓혀줘`
   - Current: `paragraph.lineSpacing = 1.40`
   - Expected action: `format_selection.lineSpacing = 1.50`
   - Must not choose `1.15` or `1.25`.

3. Korean line spacing decrease:
   - Message: `선택한 블록의 줄간격을 조금 줄여줘`
   - Current: `paragraph.lineSpacing = 1.40`
   - Expected action: `format_selection.lineSpacing = 1.30`

4. Font size increase:
   - Message: `선택한 텍스트를 조금 크게 해줘`
   - Current: `fontSize = 18`
   - Expected action: `format_selection.fontSize = 19`

5. Font size decrease:
   - Message: `선택한 텍스트를 조금 작게 해줘`
   - Current: `fontSize = 18`
   - Expected action: `format_selection.fontSize = 17`

6. Paragraph space after increase:
   - Message: `선택한 문단 뒤 간격을 조금 늘려줘`
   - Current: `paragraph.spaceAfter = 6`
   - Expected action: `format_selection.spaceAfter = 9`

7. Paragraph space after missing:
   - Message: `선택한 문단 뒤 간격을 조금 늘려줘`
   - Current: `paragraph.spaceAfter = null`
   - Expected result: no edit action; Korean explanation that current value is unavailable.

8. Unsupported table/chart/group:
   - Selection has `hasTable` or `hasChart` true, or group type.
   - Expected result: no edit action; Korean explanation that this ExecPlan only supports a selected text shape.

9. Multiple selected shapes:
   - Selection contains two or more shapes.
   - Expected result: no edit action; ask the user to select exactly one text box.

10. No selected shape:
    - Selection is empty.
    - Expected result: no edit action; ask the user to select a text box.

11. Read-only mode:
    - Same message but requested mode is review.
    - Expected result: no edit action; explanation only.

### Integration smoke tests

Use disposable PowerPoint presentations only. Do not use user business decks.

1. Context API smoke:
   - Create a disposable text box with `lineSpacing = 1.00`.
   - Verify `/api/ppt/context` returns selection or active slide shape data with `paragraph.lineSpacing`.

2. Plan API smoke:
   - Select the disposable text box.
   - Send `/api/plans` or `/api/chat/plan` with requested mode `edit` and message `선택한 블록의 줄간격을 조금 넓혀줘`.
   - Verify the returned public plan action has `lineSpacing > 1.00`.

3. Higher-current line spacing smoke:
   - Set current `lineSpacing = 1.40`.
   - Send the same request.
   - Verify returned target is `1.50`, not `1.15` or `1.25`.

4. Apply smoke:
   - Commit/apply the generated plan on a disposable deck.
   - Verify PowerPoint changed the selected shape and bridge result includes `lineSpacing` in `changedProperties`.

5. No-op smoke:
   - Reapply the same computed value or request a value already present.
   - Verify existing ExecPlan 001 no-op reporting still surfaces `변경 없음`.

6. Static prompt safety:
   - Search `server.js` and `public/taskpane.js` for old guessed-value wording.
   - Expected: no prompt instruction telling the model to guess `1.15` or `1.25`.

### Completed Test Results

All required ExecPlan 002 cases were exercised after implementation:

1. `선택한 블록의 줄간격을 조금 넓혀줘`, current `lineSpacing = 1.00` -> `lineSpacing = 1.10`: passed.
2. Same increase request, current `lineSpacing = 1.40` -> `lineSpacing = 1.50`: passed; did not choose `1.15` or `1.25`.
3. `선택한 블록의 줄간격을 조금 줄여줘`, current `lineSpacing = 1.40` -> `lineSpacing = 1.30`: passed.
4. `선택한 텍스트를 조금 크게 해줘`, current `fontSize = 18` -> `fontSize = 19`: passed.
5. `선택한 텍스트를 조금 작게 해줘`, current `fontSize = 18` -> `fontSize = 17`: passed.
6. Missing current `paragraph.lineSpacing` -> no edit action with Korean explanation: passed in direct server-helper smoke.
7. Table, chart, and group selections -> no edit action: passed.
8. Multiple selected shapes -> no edit action: passed.
9. No selected shape -> no edit action: passed.
10. Review mode -> no edit action: passed.
11. Unsupported `lineRuleWithin = 0` and unsafe current `lineSpacing = 4.00` -> no edit action: passed.
12. Existing ExecPlan 001 no-op reporting -> `noOp: true`: passed.
13. Supported paragraph-after spacing increase, current `spaceAfter = 6` -> `spaceAfter = 9`: passed as an extra scope check.

## Validation and Acceptance

Run these checks from the repository root:

```powershell
node --check server.js
node --check public\taskpane.js
git diff --check
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$null = [scriptblock]::Create((Get-Content -Raw .\scripts\ppt-bridge.ps1)); 'ppt-bridge.ps1 parse ok'"
```

If `node` is not on PATH, use the bundled Node executable recorded in `WORKLOG.md`:

```powershell
C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check server.js
C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check public\taskpane.js
```

Acceptance is met when:

- The model is still not asked to guess absolute values for relative formatting.
- The server detects supported relative selected-text formatting requests.
- The server computes exact target values from current observed values.
- Increase targets are greater than current values.
- Decrease targets are less than current values.
- Missing current values are rejected or kept as review, not guessed.
- Table, chart, grouped-shape, no-selection, and multi-selection cases are rejected safely.
- `format_selection` still freezes the plan-time selection.
- No-op formatting is still reported as `변경 없음`.
- `deckFingerprint` semantics remain unchanged and `templateFingerprint` stays separate.
- The PowerShell COM bridge remains the only write path.

## Idempotence and Recovery

These changes should be source-code edits plus ignored smoke-test files only. They can be rerun safely on disposable PowerPoint presentations.

If the compiler creates an unsafe action, disable the call to `compileRelativeSelectionFormattingPlan` in `createPlan` and return to ExecPlan 001 behavior, which safely refuses relative formatting without exact values.

Do not delete runtime user data, certificates, `.env`, copied decks, pasted images, or transaction data.

## Artifacts and Notes

Commands and validation evidence:

- `node --check server.js`: could not run because `node` is not on PATH.
- `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check server.js`: passed.
- `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check public\taskpane.js`: passed.
- `git diff --check`: passed with only the existing LF-to-CRLF working-copy warning for `server.js`.
- `powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "\`$null = [scriptblock]::Create((Get-Content -Raw .\scripts\ppt-bridge.ps1)); Write-Output 'ppt-bridge.ps1 parse ok'"`: passed.
- Static search confirmed product prompt wording still tells the model not to invent relative formatting values; `1.15` and `1.25` remain only in docs as forbidden examples.

Smoke evidence:

- Disposable PowerPoint smoke script: `C:\Users\saman\Documents\Codex\2026-06-26\read-agents-md-docs-codex-plans\work\execplan002-relative-formatting-smoke.ps1`.
- Final smoke output ended with `EXECPLAN002_RELATIVE_FORMATTING_SMOKE_OK`.
- Verified `lineSpacing` examples: `1.00 -> 1.10`, `1.40 -> 1.50`, and `1.40 -> 1.30`.
- Verified `fontSize` examples: `18 -> 19` and `18 -> 17`.
- Verified `spaceAfter` example: `6 -> 9`.
- Verified no-edit cases: review mode, unsupported `lineRuleWithin`, unsafe large line spacing, no selected shape, multiple selected shapes, table selection, chart selection, and group selection.
- Verified existing no-op reporting still returns `noOp: true`.
- Missing current line spacing was verified with a direct server-helper smoke because normal PowerPoint text boxes provide a current COM value.

Scope confirmation:

- Product code changed only in `server.js`.
- `public/taskpane.js`, `scripts/ppt-bridge.ps1`, schemas, manifest, OAuth flow, and PowerPoint write path were not changed.
- No runtime files, certs, `.env`, copied decks, pasted images, or transaction data were added.
