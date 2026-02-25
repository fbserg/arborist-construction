#!/usr/bin/env bash
# Source this file to set editing session variables:
#   source ".agents/skills/editing-arborist-reports/scripts/setup.sh"

PY="python3"
PANDOC="pandoc"

SKILL=$(find ~/.claude/plugins/cache/anthropic-agent-skills/document-skills \
    -type d -name "docx" 2>/dev/null | head -1)
SKILL_OFFICE="$SKILL/scripts/office"

PROJECT_ROOT="/home/serg/projects/arborist-construction"

echo "Setup complete."
echo "  PROJECT_ROOT : $PROJECT_ROOT"
echo "  SKILL_OFFICE : $SKILL_OFFICE"
