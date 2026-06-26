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
