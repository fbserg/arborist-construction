# Exploring Sample Reports

Load this when you need to check formatting or structure against a completed sample report.

## When to Use

- Understanding Section 5 layout for a new injury/removal type
- Matching exact table cell formatting or styling
- Checking narrative paragraph structure against a completed report

## Quick Content Review (Pandoc)

```powershell
& $PANDOC "C:\Projects\Arborism\work\[Sample] Report.docx" -t markdown
```

Sufficient for understanding content structure, section order, narrative content, and table headers.

## XML Unpack (Exact Formatting Only)

Only unpack when you need exact cell styling, border widths, font sizes, or shading values:

```powershell
& $PY $UNPACK "C:\Projects\Arborism\work\[Sample] Report.docx" "C:\Projects\Arborism\work\.sample-temp"
# Grep for relevant portions rather than reading the full file
Remove-Item -Recurse -Force "C:\Projects\Arborism\work\.sample-temp"
```

## Tips

- Prefer reusing formatting from the **current project's** XML over unpacking a sample
- Grep for known header text (e.g., "Injury source") to jump to the right section
- Never modify sample reports
