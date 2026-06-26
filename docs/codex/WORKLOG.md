# WORKLOG — Local GPT for PowerPoint Codex work

This file records what Codex changed, why, and how the result was validated. Update it at the end of every Codex task.

## 2026-06-26 — Handoff setup

Summary: Created repository-level Codex operating instructions and the first ExecPlan for safer PowerPoint harness work.

Current project memory:

- The project targets Office 2016 Professional Plus.
- The taskpane is local at `https://localhost:8765`.
- `server.js` owns local API, planner prompt, validation, transaction, preview, and commit flow.
- `scripts/ppt-bridge.ps1` is the only PowerPoint write path and uses COM automation.
- Do not add Office.js write paths or SharedRuntime requirements.
- `legacyActions` must remain for compatibility until safer primitive operations are introduced.
- The first major safety issue is not that the bridge cannot read PowerPoint. The bridge already reads a lot of rich context. The main issue is that the server and planner do not preserve or use enough of it.
- The first implementation task is `docs/codex/EXECPLAN_001_OBSERVATION_V2_PROMPT_FREEZE_NOOP.md`.

Validation evidence: Handoff files only; no repo code changed by this setup package.

## 2026-06-26 — ExecPlan 001 observation/prompt/frozen-selection/no-op implementation

Summary: Implemented `docs/codex/EXECPLAN_001_OBSERVATION_V2_PROMPT_FREEZE_NOOP.md` only. The server now preserves richer PowerPoint observation fields, sends compact template/theme/layout context to the planner, avoids prompt wording that asks the model to guess absolute line-spacing values, freezes `format_selection` targets from the plan-time selection, and marks zero-change formatting applies as no-op. The taskpane now shows `변경 없음` for no-op formatting results.

Changed files:

- `server.js`
- `public/taskpane.js`
- `docs/codex/EXECPLAN_001_OBSERVATION_V2_PROMPT_FREEZE_NOOP.md`
- `docs/codex/WORKLOG.md`

Validation evidence:

- `node --check server.js` was unavailable from PATH because `node` was not installed there.
- Bundled Node `...\node.exe --check server.js` passed.
- Bundled Node `...\node.exe --check public\taskpane.js` passed.
- `git diff --check` passed with only LF-to-CRLF working-copy warnings.
- `powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "\`$null = [scriptblock]::Create((Get-Content -Raw .\scripts\ppt-bridge.ps1)); Write-Output 'ppt-bridge.ps1 parse ok'"` passed.
- Manual live PowerPoint smoke test was not run because no `POWERPNT` process was active.

Remaining risks:

- No live COM apply was exercised in this pass.
- No-op reporting is intentionally minimal and relies on the bridge's existing `changed` counts rather than a full before/after expected-diff verifier.

## 2026-06-26 — ExecPlan 001 smoke test gate

Summary: Ran the uploaded ExecPlan 001 smoke-test prompt against the current implementation. No ExecPlan 002 work was done. No commit or push was made.

Repo state:

- Branch: `codex-handoff-001`
- Remote: `https://github.com/77875dscrj-ship-it/local-gpt-powerpoint.git`
- Pre-test uncommitted files were the ExecPlan 001 implementation files and docs: `server.js`, `public/taskpane.js`, `docs/codex/EXECPLAN_001_OBSERVATION_V2_PROMPT_FREEZE_NOOP.md`, and `docs/codex/WORKLOG.md`.

Validation evidence:

- `node --check server.js` failed from PATH because `node` is not installed there.
- `node --check public\taskpane.js` failed from PATH for the same reason.
- Bundled Node `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check server.js` passed.
- Bundled Node `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check public\taskpane.js` passed.
- `git diff --check` passed with only LF-to-CRLF working-copy warnings.
- PowerShell bridge parse-check passed with `ppt-bridge.ps1 parse ok`.

Smoke test evidence:

- Disposable bridge smoke script under ignored `runtime\smoke\` passed with `EXECPLAN001_BRIDGE_SMOKE_OK`.
- Bridge smoke verified frozen selection: A was bolded even though B was selected before apply; B did not become bold.
- Bridge smoke verified no-op base behavior: applying the same bold action a second time produced `changedTotal: 0`.
- Bridge smoke verified observation base behavior: active slide shape map included paragraph data.
- Server context API smoke passed: `templateContext`, `templateFingerprint`, `deckFingerprint`, shape `fontName`, `fontSize`, and paragraph fields including `lineSpacing` were present.
- Server apply API no-op smoke passed with `EXECPLAN001_API_NOOP_SMOKE_OK`; server response had `noOp: true`, result `changed: 0`, and a non-empty Korean no-op reason beginning with `변경 없음`.
- Static prompt safety check passed: product code no longer contains `1.15`, `1.25`, or old line-spacing guess wording.

Fix-up notes:

- No product-code fix-up was needed.
- One ignored runtime smoke script assertion was adjusted because a PowerShell regex check against Korean text failed despite the server returning the expected Korean no-op reason. The product response itself was valid.
- The local server on port 8765 was restarted so API tests used the current working tree instead of an older server process from 11:20.

Untested / remaining risks:

- Automated run did not perform the full Office taskpane/Codex OAuth UI smoke because it requires manual interaction inside the PowerPoint add-in taskpane.
- The automated smoke gate covered parser checks, bridge selection behavior, bridge no-op behavior, server context preservation, server no-op reporting, and taskpane no-op display code.

Manual UI smoke update:

- User manually selected a PowerPoint text box and asked: `선택한 블록의 줄간격을 조금 넓혀줘.`
- The taskpane did not blindly apply an arbitrary `lineSpacing` value such as `1.15` or `1.25`.
- The taskpane safely refused the relative spacing edit and explained that an exact value is needed.
- The taskpane also showed `변경 없음` for no-op formatting.
- This passes ExecPlan 001's safety goal. The remaining UX gap should be handled in ExecPlan 002 as a deterministic relative formatting compiler.

## 2026-06-26 - ExecPlan 002 deterministic relative formatting compiler

Summary: Implemented `docs/codex/EXECPLAN_002_RELATIVE_FORMATTING_COMPILER.md` only. The server now detects supported relative selected text-shape formatting requests and computes exact `format_selection` values from the current PowerPoint observation instead of asking the model to guess. The implementation covers font size, paragraph line spacing, and paragraph `spaceAfter`, while safely refusing missing values, review mode, no selection, multiple selections, tables, charts, grouped shapes, unsupported line-spacing units, and unsafe line-spacing ranges.

Changed files:

- `server.js`
- `docs/codex/EXECPLAN_002_RELATIVE_FORMATTING_COMPILER.md`
- `docs/codex/WORKLOG.md`

Implementation notes:

- Product code changed only in `server.js`.
- `public/taskpane.js` was intentionally left unchanged.
- `scripts/ppt-bridge.ps1` was intentionally left unchanged; the PowerShell COM bridge remains the only PowerPoint write path.
- No Office.js write path, SharedRuntime requirement, schema rewrite, `legacyActions` removal, Codex OAuth change, or `deckFingerprint` semantic change was added.
- Korean parser safety was implemented field-first: `줄간격` / `줄 간격` is treated as the line-spacing field, and only clear expressions such as `줄여`, `작게`, or `좁혀` are treated as decrease.
- Line spacing compilation rejects unsupported `lineRuleWithin` values and current values outside `0.80` to `3.00`; it does not clamp unsafe large values down for increase requests.
- In edit mode, the compiler can create the `format_selection` action even when the model returned review/no numeric action.

Validation evidence:

- `node --check server.js` failed from PATH because `node` is not installed there.
- Bundled Node `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check server.js` passed.
- Bundled Node `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check public\taskpane.js` passed.
- `git diff --check` passed with only the LF-to-CRLF working-copy warning for `server.js`.
- PowerShell bridge parse-check passed with `ppt-bridge.ps1 parse ok`.
- Static prompt search confirmed the product prompt still tells the model not to invent relative formatting values; `1.15` and `1.25` appear only in documentation as forbidden examples.

Smoke test evidence:

- Disposable PowerPoint smoke script used outside the repository runtime tree: `C:\Users\saman\Documents\Codex\2026-06-26\read-agents-md-docs-codex-plans\work\execplan002-relative-formatting-smoke.ps1`.
- Final smoke output ended with `EXECPLAN002_RELATIVE_FORMATTING_SMOKE_OK`.
- Verified `선택한 블록의 줄간격을 조금 넓혀줘` with current `lineSpacing = 1.00` produced target `1.10`.
- Verified the same request with current `lineSpacing = 1.40` produced target `1.50`, not `1.15` or `1.25`.
- Verified `선택한 블록의 줄간격을 조금 줄여줘` with current `lineSpacing = 1.40` produced target `1.30`.
- Verified `선택한 텍스트를 조금 크게 해줘` with current `fontSize = 18` produced target `19`.
- Verified `선택한 텍스트를 조금 작게 해줘` with current `fontSize = 18` produced target `17`.
- Verified `spaceAfter = 6` increased to `9`.
- Verified review mode, unsupported `lineRuleWithin`, unsafe `lineSpacing = 4.00`, no selected shape, multiple selected shapes, table selection, chart selection, and group selection all produced no edit action.
- Verified existing ExecPlan 001 no-op reporting still returns `noOp: true`.
- Verified missing current `paragraph.lineSpacing` with a direct server-helper smoke because normal PowerPoint text boxes provide a current COM value.

Remaining risks:

- The smoke tests used disposable decks and direct apply API calls; they did not repeat a full manual taskpane/OAuth interaction after implementation.
- Combined requests that mix a relative formatting change with unrelated edits are intentionally collapsed to the compiler-created relative `format_selection` action in this initial scope.
