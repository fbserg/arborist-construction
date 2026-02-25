# Exploring Sample Reports

Load this when you need to check formatting or structure against a completed sample report.

## When to Use

- Understanding Section 5 layout for a new injury/removal type
- Matching exact table cell formatting or styling
- Checking narrative paragraph structure against a completed report

## Quick Content Review (Pandoc)

```bash
pandoc "$PROJECT_ROOT/work/[Sample] Report.docx" -t markdown
```

Sufficient for understanding content structure, section order, narrative content, and table headers.

## XML Unpack (Exact Formatting Only)

Only unpack when you need exact cell styling, border widths, font sizes, or shading values:

```bash
cd "$SKILL_OFFICE" && python3 unpack.py "$PROJECT_ROOT/work/[Sample] Report.docx" "$PROJECT_ROOT/work/.sample-temp"
# Grep for relevant portions rather than reading the full file
rm -rf "$PROJECT_ROOT/work/.sample-temp"   # always clean up
```

## Tips

- Prefer reusing formatting from the **current project's** XML over unpacking a sample
- Grep for known header text (e.g., "Injury source") to jump to the right section
- Never modify sample reports
