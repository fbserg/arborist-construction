# Integration Test: Client Revision Cycle

A recurring end-to-end test that exercises the full edit workflow cold:
1. Client email → Claude asks discovery questions
2. Client answers → edit script covers all downstream changes
3. Output verified with pandoc acceptance test

**Run frequency:** Each new client revision. Keep test fresh by randomizing tree targets.

## How to Run

1. Place a fresh report in `new/`
2. Paste the client email from `scenario.md` as the starting prompt
3. Answer Claude's discovery questions using `scenario.md` (or generate fresh answers)
4. Claude executes the edit script and verifies output
5. Review `complete/` output in Word — accept all changes, no repair warnings

## Randomization (per run)

At the top of each generated `edit_script.py`, record the run's choices:

```python
# === TEST RUN CONSTANTS ===
REPORT        = "73 Kennedy Avenue Report.docx"
REMOVAL_INJURY_TREE   = 3   # tree: removal → injury
CONFIRMED_REMOVAL     = 2   # tree: confirmed removal
CONDITION_REMOVAL     = 8   # tree: condition-based removal
NEW_TREE_NUMBER       = 9   # tree to insert
INJURY_PROFILE        = "A" # A/B/C/E
INJURY_DIRECTION      = "South"
CONDITION_CAUSE       = "deadwood + co-dominant"
RSE_EXPANSION         = "standard"  # standard / site-specific
NEW_TREE_PROFILE      = "C"
NEW_TREE_DIRECTION    = "North"
```

Vary these each quarter to prevent muscle memory:
- Swap which trees are targets
- Rotate injury profile (A/B/C/E)
- Vary condition cause (deadwood / lean / decline / cavities)
- Vary RSE expansion depth (1 sentence vs full protocol paragraph)
- Vary new tree species and profile

## Verification Checklist

- [ ] pandoc `--track-changes=accept` grep confirms all new text present
- [ ] pandoc `--track-changes=all` shows del/ins pairs (not just insertions)
- [ ] Open in Word → Accept All → no repair warnings
- [ ] Summary section counts updated (removals, permits)
- [ ] Cross-section consistency: Summary ↔ impact table ↔ Section 5 narrative
- [ ] Changelog written with change ID range and RSID

## Watch List

| Risk | What to check |
|------|---------------|
| Table row tracked-change insertion | Word repair warnings; malformed `<w:trPr><w:ins>` |
| Multi-para deletion (removal block) | Wrong paraId anchor; orphaned XML |
| Mini table generated from scratch | Column widths/borders match existing |
| Summary count cascade | All 4 counts in top table updated |
| `replace_phrase_in_run` on multi-run spans | Fails silently — verify with `find_run()` first |
