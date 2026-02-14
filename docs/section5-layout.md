# Section 5 (Conclusion) Layout

Read this when editing Section 5 of a report. For overview, see `CLAUDE.md`.

## Structure Order

Do not reorder or skip structural elements:

1. **Permit summary** — Bullet list of removal/injury permit counts
2. **Protection note** — Standard paragraph re: TPZ fencing
3. **No other trees note** — Standard paragraph
4. **Impact summary table** — Tree #, Source of impact, Direction (one row per tree, multi-row sources merged)
5. **"Injuries" subheading**
   - Per injured tree:
     a. **Mini data table** — TREE #, Species, DBH, Condition, Ownership, Direction, TPZ, % Encroachment, Permit
     b. **Injury detail table** — Injury source, Closest point of impact, Max depth, Impact to condition (one row per injury source)
     c. **Narrative paragraphs** — Written per `guideline.md` Template 1
   - After all injury narratives: RSE boilerplate paragraphs
6. **"Removals" subheading** (if applicable)
   - Per removed tree: narrative per Template 2 or 3
7. **"General Notes"** — Standard closing paragraphs
8. **Signature block**

Injury narratives go **after** the injury detail table, **before** the RSE boilerplate. They do NOT go after the RSE boilerplate or in the empty paragraphs between RSE and Removals.

## Standard XML Anchors

Use with `editor.get_node` to locate insertion points:

| Anchor text | Locates |
|---|---|
| `"Injuries"` | Injuries subheading paragraph |
| `"Injury source"` | Injury detail table header row |
| `"Excavation should not be deeper"` | RSE boilerplate (insert narrative before this) |
| `"Removals"` | Removals subheading |
