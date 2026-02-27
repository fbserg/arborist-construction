# Setup

## Prerequisites

- **Python 3.12+**
- **pandoc** (3.1+)
- **pip packages**: `defusedxml`, `openpyxl`

## macOS

```bash
brew install python3 pandoc
pip3 install defusedxml openpyxl
```

Set `PROJECT_ROOT` if your checkout isn't at the default WSL path:

```bash
export PROJECT_ROOT="$HOME/projects/arborist-construction"
```

Find the Claude Code docx plugin path (needed for pack/unpack):

```bash
find ~/.claude -path "*/document-skills/*/skills/docx" -type d
```

Update the `Current hash` in CLAUDE.md if it differs from your install.

## WSL / Linux (current primary environment)

```bash
sudo apt install python3 pandoc
pip3 install defusedxml openpyxl
```

Default project root: `/home/serg/projects/arborist-construction`
Default plugin hash: `1ed29a03dc85`

## Verify

```bash
python3 --version    # 3.12+
pandoc --version     # 3.1+
python3 -c "import defusedxml; import openpyxl; print('ok')"
```
