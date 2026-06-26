# ExecPlan 001 — Connect Observation v2 to planner, stop absolute-value guessing, fix frozen selection timing, and report no-op formatting

This ExecPlan follows `docs/codex/PLANS.md`. Keep the `Progress`, `Surprises & Discoveries`, `Decision Log`, and `Outcomes & Retrospective` sections updated as work proceeds.

## Purpose / Big Picture

Today the add-in can read and edit PowerPoint, but the editing harness is not safe enough. The PowerShell bridge already reads useful shape/style/template information, but the Node server drops much of that information before the planner sees it. The planner is also encouraged to guess absolute values such as `lineSpacing: 1.15`, which can accidentally reduce spacing when the current value is already larger. In addition, `format_selection` can be compiled again at preview or commit time using the currently active selection instead of the selection that was active when the plan was created.

After this change, Codex/ChatGPT should receive richer read-only context, should no longer be instructed to guess absolute formatting values for relative requests, should apply `format_selection` to the plan-time selected shape rather than a later accidental selection, and should clearly report when a formatting action made no actual change.

This is not a full primitive-operation rewrite. It is a small harness safety step that preserves the current Office 2016 + COM bridge architecture.

## Progress

- [x] (2026-06-26 15:06 local) Initial investigation of `server.js`, `scripts/ppt-bridge.ps1`, schemas, and taskpane result reporting is complete.
- [x] (2026-06-26 15:09 local) Richer shape/style/paragraph fields are preserved through `normalizeShape`, `augmentSelection`, and `shapeMapForPrompt`.
- [x] (2026-06-26 15:09 local) Compact template context and separate `templateFingerprint` are added without changing `deckFingerprint` semantics.
- [x] (2026-06-26 15:10 local) Planner prompt no longer tells the model to guess `lineSpacing` values such as `1.15` or `1.25` for relative requests.
- [x] (2026-06-26 15:10 local) `format_selection` target freezing uses the plan creation context, not the commit-time active selection.
- [x] (2026-06-26 15:11 local) Minimal no-op reporting is added for formatting actions where all target changes report `changed = 0`.
- [x] (2026-06-26 15:14 local) Validation commands and any manual test notes are recorded.

## Surprises & Discoveries

- The current Codex thread folder was empty except for `work/` and `outputs/`; the requested repo files were found at `C:\Users\saman\Documents\Codex\2026-06-25\gpt-for-powerpoint\outputs\local-gpt-powerpoint`.
- `node` was not available on PATH in this shell. Validation used the bundled Node executable at `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe`.
- `scripts/ppt-bridge.ps1` already returned font, paragraph, layout, design, and theme information. The server was the place dropping most of that read-only context before prompting.
- `previewPlan` and `directCommitPlan` recompiled plans with the current context. That meant `format_selection` could refreeze whichever shape was active at preview or commit time.

## Decision Log

- Decision: Keep `legacyActions` as the execution adapter for this task.
  Rationale: The immediate goal is safer observation and target handling, not a full schema migration. Removing `legacyActions` would risk breaking existing behavior.
  Date/Author: 2026-06-26 / planning handoff

- Decision: Keep `deckFingerprint` unchanged and add `templateFingerprint` separately.
  Rationale: `deckFingerprint` is used for edit-conflict detection. Mixing template/master/theme metadata into it could make old transactions conflict for reasons unrelated to slide content changes.
  Date/Author: 2026-06-26 / planning handoff

- Decision: The planner should not guess absolute values for relative formatting requests.
  Rationale: A request like "increase spacing" is meaningful only relative to the current value. Guessing `1.15` can be wrong if the current value is already higher.
  Date/Author: 2026-06-26 / planning handoff

- Decision: Recompile selected actions from `planRecord.context` in preview/direct commit.
  Rationale: `planRecord.context` is the snapshot captured when the plan was created. Current context is still used for conflict checks, but not for selecting a new target shape.
  Date/Author: 2026-06-26 / Codex

- Decision: Keep no-op detection minimal and server-led.
  Rationale: The bridge already reports `changed` per formatting target. Marking zero-change `format_selection` results with `noOp` avoids a larger expected-diff verifier in this task.
  Date/Author: 2026-06-26 / Codex

## Outcomes & Retrospective

Implementation is complete. `server.js` now preserves richer shape/style/paragraph observation data, passes compact template/theme/layout context to the planner with a separate `templateFingerprint`, removes prompt wording that encouraged guessed line-spacing values, freezes `format_selection` from plan creation context during preview/direct commit, and marks zero-change formatting results as no-op. `public/taskpane.js` now displays `변경 없음` for those no-op formatting results.

Validation passed for server syntax, taskpane syntax, diff whitespace, and PowerShell bridge parsing. The live PowerPoint smoke test was not run because no `POWERPNT` process was active in this session; that test still needs a running PowerPoint 2016 instance and a selected text box.

Remaining risks: no live COM apply was exercised in this pass, and no-op detection is intentionally minimal. It trusts the bridge's existing `changed` counts and does not compare an expected before/after property diff.

## Context and Orientation

This repository is a local PowerPoint 2016 add-in. PowerPoint shows a local taskpane from `public/taskpane.html` and `public/taskpane.js`. The taskpane talks to the local HTTPS server in `server.js`. The server builds a prompt for Codex/ChatGPT OAuth, receives a strict JSON presentation plan, validates it, compiles it into an execution plan, and applies it through `scripts/ppt-bridge.ps1`. The bridge uses PowerPoint COM automation, which means it controls the running PowerPoint application through Windows automation objects.

The important current flow is:

1. `createPlan` in `server.js` calls `getContextSnapshot`.
2. `getContextSnapshot` runs the bridge action `context`.
3. `scripts/ppt-bridge.ps1` returns deck, slide, active slide shape map, selection, and template data.
4. `augmentContext`, `normalizeShape`, `augmentSelection`, and `shapeMapForPrompt` prepare data for storage and prompting.
5. `buildPlannerPrompt` sends a compact PowerPoint snapshot to the model.
6. `normalizePresentationPlan` and `validatePresentationPlan` validate the model output.
7. `compileExecutionPlan` converts `legacyActions` to execution operations and currently freezes `format_selection` targets from the context passed into it.
8. `previewPlan`, `directCommitPlan`, and `commitTransaction` may read current context again.
9. `apply-json` in `scripts/ppt-bridge.ps1` applies the plan, and `public/taskpane.js` displays the result.

The current bridge already reads rich shape data in `Get-ShapeInfo`, including font, paragraph, table/chart, fill/line, tags, and placeholder information. The server currently keeps only part of that data. This task should preserve more of it without changing the bridge write path.

## Plan of Work

First, preserve richer observation fields in `server.js`. Update `normalizeShape` so it retains `fontName`, `bold`, `fontRgb`, a nested `paragraph` object, and any other existing read-only fields that are already emitted by `scripts/ppt-bridge.ps1`. Update `shapeMapForPrompt` so the planner sees a compact version of these fields for active slide shapes. Make sure `selection` also retains these fields through `augmentSelection`.

Second, add compact template context to the planner prompt. The bridge already returns `template` from `Get-Context`. In `augmentContext`, compute a separate `templateFingerprint` from a compact, stable subset of template data. Do not include template information in the existing `deckFingerprint` basis. In `buildPlannerPrompt`, include a compact `templateContext` with theme name, heading/body fonts, used layouts, used designs, and active slide layout/design names if available.

Third, revise the planner prompt. In `buildPlannerPrompt`, remove language that tells the model to set line spacing to `1.15` or `1.25` or paragraph gaps to fixed guessed values for relative requests. Replace it with guidance that relative requests should be represented as intent and direction in `assistantMessage`, `intent`, `outline`, or warnings unless the user explicitly gives an exact value. Because the current schema still uses `legacyActions`, keep compatibility, but do not encourage invented absolute values.

Fourth, fix `format_selection` target timing. The selected shape should be the shape selected when the plan was created. The current active selection at preview or commit time should be used only to verify that the deck has not changed unexpectedly, not as the target source. The simplest safe approach is to ensure preview and direct commit compile `format_selection` actions using `planRecord.context`, then separately compare affected slide fingerprints against the current context. If `compileExecutionPlan` needs both a target-freeze context and a current-validation context, add a small options object rather than overloading one context.

Fifth, add minimal no-op reporting. The bridge already reports `changed` and `changedProperties` for formatting targets. In `server.js`, after apply, detect the case where an execution contains one or more `format_selection` actions and every corresponding result has `changed` equal to zero. In that case, add `noOp: true` and a clear reason to the result. In `public/taskpane.js`, display "변경 없음" instead of a normal success line when `noOp` is true. This is a minimal first step; do not build a full expectedDiff verifier in this task.

## Concrete Steps

Work from the repository root.

1. Inspect the current files:

    git status --short
    node --check server.js

2. Edit `server.js`:

    - Update `normalizeShape`.
    - Update `shapeMapForPrompt`.
    - Add compact template context and `templateFingerprint` in `augmentContext` or a small helper function.
    - Update `buildPlannerPrompt` to include compact template context and remove guessed absolute spacing guidance.
    - Update `compileExecutionPlan` or its callers so `format_selection` freezes targets from plan creation context.
    - Add minimal no-op detection after bridge apply.

3. Edit `public/taskpane.js`:

    - When the server result has `noOp: true`, show a Korean message beginning with `변경 없음:`.
    - Keep existing success output for real changes.

4. Do not edit `manifest.xml` unless a validation command reveals a syntax problem unrelated to this task.

5. Do not change `scripts/ppt-bridge.ps1` unless a small adjustment is strictly required to support returned data that is already being read. If changed, keep it read-only for observation or minimal result reporting.

## Validation and Acceptance

Run these checks from the repository root:

    node --check server.js
    git diff --check

On Windows, parse-check PowerShell:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$null = [scriptblock]::Create((Get-Content -Raw .\scripts\ppt-bridge.ps1)); 'ppt-bridge.ps1 parse ok'"

If PowerPoint 2016 is available, do a manual smoke test:

1. Start the local server with `start-local-gpt.cmd`.
2. Open the add-in deck.
3. Select text box A.
4. Ask for a small selected-block formatting change in edit mode.
5. After the plan appears, click a different shape B before applying.
6. Apply the change.
7. Verify that A, not B, was the target.
8. Ask for a change that is already true, such as bolding already bold text or setting a value already applied if the UI supports it.
9. Verify the taskpane reports `변경 없음` rather than a normal completion.

Acceptance is met when:

- The planner prompt contains rich style/paragraph/template context.
- The prompt no longer instructs the model to guess specific line-spacing numbers for relative requests.
- `format_selection` targets the plan-time selection.
- A no-op formatting apply is clearly reported as no-op.
- The existing COM bridge write path remains the only write path.
- `deckFingerprint` behavior for slide/content conflict detection is not expanded to include template context.

## Idempotence and Recovery

The changes are source-code edits only. They can be retried safely from a clean git branch. If a change breaks planning or commit flow, revert the specific file with `git checkout -- <file>` or use your normal Git rollback process. Do not delete runtime user data, certificates, or local environment files.

## Artifacts and Notes

Implementation notes:

- `server.js` now keeps `fontName`, `bold`, `fontRgb`, and nested paragraph spacing/rule fields on normalized shapes. Because `augmentSelection` uses `normalizeShape`, selected shapes keep the same richer fields.
- `server.js` now computes `templateContext` and `templateFingerprint` separately from `deckFingerprint`; the deck fingerprint basis was not expanded.
- The planner prompt now includes `templateContext` and `templateFingerprint`, and no longer asks the model to invent fixed line-spacing values for relative spacing requests.
- `previewPlan` and `directCommitPlan` now call `compileExecutionPlan(..., planRecord.context, ...)`; current context remains part of affected-slide conflict checks.
- `public/taskpane.js` now displays `변경 없음` for no-op `format_selection` results.

Validation evidence:

- `node --check server.js` could not run from PATH because `node` is not installed there.
- Bundled Node: `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check server.js` passed.
- Bundled Node: `C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe --check public\taskpane.js` passed.
- `git diff --check` passed. Git printed only LF-to-CRLF working-copy warnings.
- `powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "\`$null = [scriptblock]::Create((Get-Content -Raw .\scripts\ppt-bridge.ps1)); Write-Output 'ppt-bridge.ps1 parse ok'"` passed with `ppt-bridge.ps1 parse ok`.
- `Get-Process POWERPNT -ErrorAction SilentlyContinue` found no running PowerPoint process, so the manual add-in smoke test was not run.

Smoke test evidence added 2026-06-26 16:00 local:

- Repo state check passed: branch `codex-handoff-001`; remote `https://github.com/77875dscrj-ship-it/local-gpt-powerpoint.git`.
- `node --check server.js` and `node --check public\taskpane.js` still fail from PATH because `node` is not installed there; bundled Node equivalents passed.
- `git diff --check` passed with only LF-to-CRLF working-copy warnings.
- `scripts\ppt-bridge.ps1` parse-check passed.
- Disposable COM bridge smoke passed with `EXECPLAN001_BRIDGE_SMOKE_OK`: frozen selection targeted A even after B was selected before apply, and the second identical apply reported `changedTotal: 0`.
- Server context API smoke passed: normalized context included `templateContext`, `templateFingerprint`, `deckFingerprint`, shape `fontName`, shape `fontSize`, and paragraph fields including `lineSpacing`.
- Server apply API no-op smoke passed with `EXECPLAN001_API_NOOP_SMOKE_OK`: result contained `noOp: true`, target result `changed: 0`, and a non-empty Korean no-op reason beginning with `변경 없음`.
- Prompt safety static check passed: `server.js` and `public\taskpane.js` no longer contain `1.15`, `1.25`, or the old line-spacing guess wording.
- Full Office taskpane/Codex OAuth UI smoke was not run. It requires manual interaction inside the PowerPoint add-in taskpane; the automated pass covered code parse, bridge behavior, server context, server no-op reporting, and taskpane no-op display code.

Manual UI smoke evidence added 2026-06-26 16:23 local:

- User manually selected a PowerPoint text box and asked: `선택한 블록의 줄간격을 조금 넓혀줘.`
- The taskpane did not blindly apply an arbitrary `lineSpacing` value such as `1.15` or `1.25`.
- The taskpane safely refused the relative spacing edit and explained that an exact value is needed.
- The taskpane also showed `변경 없음` for no-op formatting.
- This is acceptable for ExecPlan 001 because this plan is about safety, observation preservation, prompt safety, frozen selection, and no-op reporting. The remaining UX gap belongs in ExecPlan 002 as a deterministic relative formatting compiler.
