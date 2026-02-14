# Exploring Sample Reports

This guide covers how to efficiently explore sample reports placed in `work/` for reference.

## When to Use Sample Reports

- Understanding Section 5 layout for a new injury/removal type
- Matching exact table cell formatting or styling
- Checking narrative paragraph structure against a completed report

## Quick Content Review (Pandoc)

Use pandoc to read the report as markdown — sufficient for understanding content structure and layout order:

```powershell
& "C:\Users\User\AppData\Local\Pandoc\pandoc.exe" "C:\Projects\Arborism\work\[Sample] Report.docx" -t markdown
```

This shows all text, tables, and headings. Use this for:
- Identifying section order (injuries, removals, general notes)
- Reading narrative paragraph content
- Checking data table column headers and values

## XML Unpack (Exact Formatting Only)

Only unpack to XML when you need exact cell styling, border widths, font sizes, or shading values. Unpack to a temporary location and clean up after:

```powershell
# Unpack sample to a temp location
& "C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe" "C:\Users\User\.claude\plugins\cache\anthropic-agent-skills\document-skills\69c0b1a06741\skills\docx\ooxml\scripts\unpack.py" "C:\Projects\Arborism\work\[Sample] Report.docx" "C:\Projects\Arborism\work\.sample-temp"

# Read the specific XML section you need
# (use grep to find relevant portions rather than reading the full file)

# Clean up when done
Remove-Item -Recurse -Force "C:\Projects\Arborism\work\.sample-temp"
```

## Tips

- Prefer reusing formatting from the **current project's** XML over unpacking a sample. The project document already has the correct styles.
- If you only need one table's cell formatting, grep for a known header text (e.g., "Injury source") in the sample XML to jump to the right section.
- Never modify sample reports. They are read-only references.
