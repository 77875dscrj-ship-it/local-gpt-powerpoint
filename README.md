# Local GPT for PowerPoint

Local GPT for PowerPoint is a local Office add-in experiment for PowerPoint 2016.

The goal is to let an AI assistant read the open PowerPoint deck, understand the selected slide or shape, create an editable plan, and apply approved changes directly to native PowerPoint objects.

## Why this exists

Modern ChatGPT/Claude PowerPoint add-ins depend on newer Office web add-in runtimes. On this machine, Office 2016 can sideload a taskpane, but newer hosted add-in web apps do not reliably run inside PowerPoint.

This project works around that by using:

- a local HTTPS taskpane at `https://localhost:8765`
- a local Node.js server
- Codex/ChatGPT OAuth, with no OpenAI API key stored in the project
- a PowerShell COM bridge that edits the active PowerPoint presentation

## Main Pieces

- `server.js` - local HTTPS API server, planner, validation, and transaction flow
- `public/` - PowerPoint taskpane UI
- `scripts/ppt-bridge.ps1` - PowerPoint COM automation bridge
- `scripts/read-clipboard-image.ps1` - Windows clipboard image fallback
- `scripts/ensure-cert.ps1` - local HTTPS certificate generator
- `scripts/make-pptx.py` - creates the sideloadable PowerPoint add-in deck
- `schemas/` - JSON schemas for model-produced presentation plans
- `config/policies.json` - edit, preview, backup, and risk policies

## Current Status

Working:

- PowerPoint 2016 taskpane launches
- local server and taskpane communicate
- Codex/ChatGPT OAuth model calls work without API keys
- current deck/slide/selection context can be read
- selected shapes can be formatted through PowerPoint COM
- pasted images can be attached through file upload or Windows clipboard fallback
- undo grouping is attempted with PowerPoint `StartNewUndoEntry`

In progress:

- stronger PowerPoint harness/context
- template, layout, master, and theme awareness
- relative formatting such as "increase line spacing" based on current values
- no-op detection and before/after property reporting
- safer primitive operations beyond the current legacy action recipes

## Not Committed

The repository intentionally excludes:

- `.env`
- local HTTPS certificates under `certs/`
- runtime plans, previews, pasted images, copied decks, and transactions under `runtime/`

Those files may contain local machine state or user presentation data.

## Start

Run:

```powershell
.\start-local-gpt.cmd
```

Then open `Local-GPT-PowerPoint.pptx` or rerun:

```powershell
.\install-and-open.ps1
```
