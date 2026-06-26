# AGENTS.md — Local GPT for PowerPoint

These are repository-level instructions for Codex. They are intentionally short, practical, and enforce the current architecture. If a task-specific ExecPlan exists under `docs/codex/`, read it before editing code.

## Project purpose

This repository is a local PowerPoint 2016 add-in experiment. The user works in Office 2016 Professional Plus, so modern Office.js/SharedRuntime write paths are not the target. The add-in shows a local taskpane, talks to a local Node.js server, asks Codex/ChatGPT OAuth for a JSON plan, validates that plan, and applies approved changes to native PowerPoint objects through the PowerShell COM bridge at `scripts/ppt-bridge.ps1`.

## Non-negotiable architecture rules

- Keep the PowerShell COM bridge as the only PowerPoint write path.
- Do not add an Office.js write path.
- Do not add SharedRuntime requirements to `manifest.xml`.
- Do not remove `legacyActions` yet. Treat them as a compatibility adapter while safer primitive operations are developed.
- Do not mix template/master/theme data into the existing `deckFingerprint`. Add a separate `templateFingerprint` for template/layout/theme data.
- Do not store or commit secrets, local certificates, `.env`, runtime presentations, copied decks, pasted images, or transaction data.
- Do not rewrite the whole architecture in one task. Prefer small, reviewable changes.

## Important files

- `README.md`: project overview and current status.
- `manifest.xml`: Office 2016 taskpane manifest.
- `server.js`: local HTTPS API server, planner prompt, plan validation, transaction/preview/commit flow.
- `scripts/ppt-bridge.ps1`: PowerPoint COM observation and editing bridge.
- `schemas/presentation-plan.schema.json`: schema for model-produced plans.
- `schemas/execution-plan.schema.json`: schema for compiled execution plans.
- `config/policies.json`: edit, preview, backup, and selection-edit policies.
- `public/taskpane.js`: taskpane UI and apply/preview result display.

## How Codex should work in this repo

1. Start each task by reading this file, `README.md`, and any task-specific file under `docs/codex/` named in the prompt.
2. For complex tasks, plan first. Do not start coding until the plan names the exact files and functions to change.
3. Keep the scope narrow. One Codex thread should handle one coherent unit of work.
4. Before editing, identify the current behavior and the user-visible problem.
5. Prefer additive changes that preserve existing behavior.
6. After editing, run the smallest relevant checks available in the current environment.
7. Update the task ExecPlan or `docs/codex/WORKLOG.md` with progress, decisions, surprises, and validation evidence.
8. Report exactly what changed, how it was validated, and what remains.

## Validation commands

Run what is available in the current environment. If a command cannot run because PowerPoint or Windows PowerShell is unavailable, record that clearly.

From the repository root:

    node --check server.js
    git diff --check

On Windows, also parse-check the PowerShell bridge without requiring a live PowerPoint instance:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$null = [scriptblock]::Create((Get-Content -Raw .\scripts\ppt-bridge.ps1)); 'ppt-bridge.ps1 parse ok'"

When a live PowerPoint instance is available, manually smoke-test the add-in flow:

    .\start-local-gpt.cmd

Then open the add-in deck, select a text box, ask for a small formatting change, apply it, and verify that the taskpane reports the actual changed properties or reports "변경 없음" when nothing changed.

## Definition of done

A task is done only when all of these are true:

- The requested behavior is implemented in the smallest reasonable scope.
- The existing Office 2016 + COM bridge architecture remains intact.
- Relevant checks were run or explicitly marked unavailable.
- The final report includes changed files, validation evidence, and remaining risks.
- Any task-specific ExecPlan or worklog has been updated.
