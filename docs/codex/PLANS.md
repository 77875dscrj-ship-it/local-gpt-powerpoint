# PLANS.md — Local GPT for PowerPoint ExecPlan rules

An ExecPlan is a self-contained Markdown work plan that Codex can follow without relying on prior chat memory. Use an ExecPlan for changes that touch multiple files, modify the editing harness, or affect PowerPoint behavior.

## Required properties

Every ExecPlan must be self-contained. A future Codex thread or human reader should be able to read only the current repo plus the ExecPlan and understand the goal, the files involved, the exact work to do, and how to prove that it works.

Every ExecPlan must be a living document. As Codex works, it must update the Progress, Surprises & Discoveries, Decision Log, and Outcomes & Retrospective sections. These sections are required.

Every ExecPlan must be outcome-focused. Do not define success as merely "code was changed". Define what a user can observe after the change and how to verify it.

Every technical term must be explained in plain language the first time it appears.

## Required sections

Use this structure for every task-specific plan:

# <Short task title>

This ExecPlan follows `docs/codex/PLANS.md`. Keep the `Progress`, `Surprises & Discoveries`, `Decision Log`, and `Outcomes & Retrospective` sections updated as work proceeds.

## Purpose / Big Picture

Explain what problem the user experiences today, what will be possible after the change, and how someone can see it working.

## Progress

Use timestamped checkboxes. Update this section at every stopping point.

- [ ] (YYYY-MM-DD HH:MM local) Initial investigation complete.
- [ ] Implementation step complete.
- [ ] Validation complete.

## Surprises & Discoveries

Record unexpected code behavior, missing tools, test limitations, or design facts found while working. Include short evidence where possible.

## Decision Log

Record decisions in this format:

- Decision: ...
  Rationale: ...
  Date/Author: ...

## Outcomes & Retrospective

At completion or at a major milestone, summarize what changed, what was validated, what remains risky, and what should be handled in a later task.

## Context and Orientation

Explain the relevant repo structure as if the reader knows nothing. Name exact files and functions.

## Plan of Work

Describe the work as small, ordered edits. Each edit should name the file and function or module.

## Concrete Steps

List exact commands to run, expected outputs, and fallback notes if a tool is unavailable.

## Validation and Acceptance

Describe how to prove the change works. Prefer behavior that a person can observe, not only internal implementation facts.

## Idempotence and Recovery

Explain how to rerun the steps safely and how to back out if something goes wrong.

## Artifacts and Notes

Keep concise evidence: command output, small diffs, manual test notes, and links to related worklog entries.
