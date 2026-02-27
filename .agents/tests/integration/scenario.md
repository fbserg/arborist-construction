# Integration Test Scenario

## Phase 1 — Client Email (paste this as the starting prompt)

---

> Subject: Revisions — [Client] Report
>
> Hi, scope has changed a bit based on latest contractor plans:
>
> - **Tree [REMOVAL_INJURY]** was listed as a removal but they're going to work around it now. It'll be impacted by the new driveway instead — interlock, the closest point is about 2.4m from the trunk.
> - **Tree [CONFIRMED_REMOVAL]** definitely needs to come out — house footprint is sitting right on top of it.
> - **Tree [CONDITION_REMOVAL]** — the homeowner wants it removed. The tree isn't in great shape honestly.
> - The RSE section for **Tree [REMOVAL_INJURY]** should be more detailed — contractor asked about the hand-digging protocol specifically.
> - I've attached an updated tree inventory, one tree was missed in the original survey. See the table below:
>
> | Tree # | Species | DBH (cm) | Condition | TPZ (m) | Ownership | Permit |
> |--------|---------|----------|-----------|---------|-----------|--------|
> | [NEW_TREE] | Silver Maple | 28 | Fair | 5.6 | Private | Yes |
>
> The rest of the tree data is unchanged.

---

*Before running: replace `[REMOVAL_INJURY]`, `[CONFIRMED_REMOVAL]`, `[CONDITION_REMOVAL]`, `[NEW_TREE]` with actual tree numbers from the report.*

---

## Phase 2 — Discovery Questions (Claude asks these)

**On the removal→injury tree:**
- Q1: What's the proposed construction — just the driveway, or anything else within the TPZ?
- Q2: What direction does the driveway approach from relative to the trunk?
- Q3: Is this tree currently in the impact summary table as a removal?

**On the confirmed removal:**
- Q4: Is this tree currently in the report (impact table, narrative), or is it new?
- Q5: Does it require a permit?

**On the condition-based removal:**
- Q6: What specific condition issues — deadwood, structural lean, cavities, or general decline?
- Q7: Does it require a permit?

**On RSE detail:**
- Q8: Standard RSE boilerplate expansion, or something site-specific?

**On the new tree:**
- Q9: What section of the site, and what's the source of impact / construction activity near it?
- Q10: What's the impact direction?

**Mandatory:** Confirm the RSID to use (from unpack output).

---

## Phase 3 — Canned Answers (for automated test run)

- A1: Just the interlock driveway. Surface hardscaping only, max 150mm depth → **Profile A**
- A2: Driveway approaches from the **south**
- A3: *(check actual report — update this field per run)*
- A4: *(check actual report — update this field per run)*
- A5: Yes, requires a permit
- A6: Deadwood throughout crown, co-dominant stems with included bark
- A7: Yes, requires a permit
- A8: Standard protocol — hand dig + AirSpade within TPZ, retain roots ≥5cm, clean cut + wound compound on smaller roots
- A9: Near the rear garage — impacted by new garage foundation. Profile C.
- A10: Impact from the **north**

---

## Phase 4 — Verification Greps (run after pack)

```bash
pandoc --track-changes=accept "$PROJECT_ROOT/complete/[Report].docx" -t plain \
    | grep -iE "(interlock|Profile A|AirSpade|deadwood|Silver Maple|[NEW_TREE_SPECIES])"
```

Expected: each keyword appears in the accepted output.
