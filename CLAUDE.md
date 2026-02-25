# CLAUDE.md

Be concise; do not explain code unless asked.

## Purpose

This repository contains arborist consulting project files. Claude instances working here create and edit arborist reports in Word (.docx) format.

## Report Work

For any report reading, editing, or creation, read **`.agents/skills/editing-arborist-reports/SKILL.md`** directly (use the `Read` tool — this is a local project skill, not a system skill). It contains the full workflow, tool paths, editing principles, and reference files.

Content rules (impact profiles, narrative templates, tone constraints) live in `guideline.md` at the project root — load it when writing or editing report narrative content.

## Version Control

This repo is pushed to GitHub. **Only guidelines, documentation, and skill definitions are tracked.** All binary/work files live on disk but are gitignored.

**Tracked:** `CLAUDE.md`, `guideline.md`, `.agents/`, `.gitignore`
**Ignored:** `*.docx`, `*.xlsx`, `*.ai`, `*.pdf`, `*.emf`, `new/`, `work/`, `complete/`, `.claude/`

Folder convention: `new/` (incoming) → `work/[Client]/` (active, with `.work/` artifacts) → `complete/` (delivered)

## Dependencies

- **Python**: `python3` (WSL, v3.12)
- **pandoc**: `pandoc` (WSL, v3.1.3)
- **defusedxml** (pip installed in WSL)

## Environment

Claude Code runs in WSL. Project root: `/home/serg/projects/arborist-construction`.
All docx tooling runs in WSL using `python3` and `pandoc` directly.

Docx plugin location: `~/.claude/plugins/cache/anthropic-agent-skills/document-skills/<hash>/skills/docx/`
Current hash: `1ed29a03dc85`
Pack/unpack scripts: `scripts/office/pack.py`, `scripts/office/unpack.py`
These require running from the `scripts/office/` directory (relative imports).
