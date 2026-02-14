# CLAUDE.md

Be concise; do not explain code unless asked.

## Purpose

This repository contains arborist consulting project files. Claude instances working here create and edit arborist reports in Word (.docx) format.

## Report Work

For any report reading, editing, or creation, use the **editing-arborist-reports** skill (`.agents/skills/editing-arborist-reports/`). It contains the full workflow, tool paths, editing principles, and reference files.

Content rules (impact profiles, narrative templates, tone constraints) live in `guideline.md` at the project root — load it when writing or editing report narrative content.

## Version Control

This repo is pushed to GitHub. **Only guidelines, documentation, and skill definitions are tracked.** All binary/work files live on disk but are gitignored.

**Tracked:** `CLAUDE.md`, `guideline.md`, `.agents/`, `work/sample-reports.md`, `.gitignore`
**Ignored:** `*.docx`, `*.xlsx`, `*.ai`, `*.pdf`, `*.emf`, `temp/`, `client/`, `.claude/`

## Dependencies

- **Python 3.12**: `C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe`
- **pandoc**: `C:\Users\User\AppData\Local\Pandoc\pandoc.exe`
- **defusedxml** and **lxml** (pip installed)
