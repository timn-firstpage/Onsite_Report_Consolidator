# Manifest and mapping contract

## Manifest

Create `run.json` in or near the selected run folder. This is a concrete example, not permission to search for matching files:

```json
{
  "run_id": "client-run-1",
  "run_folder": ".",
  "company": "Example Co",
  "site_hosts": ["www.example.com"],
  "asset_hosts": ["cdn.shopify.com"],
  "output_date": "2026-10-15",
  "sources": [
    {
      "file": "11.1-11.9_Example Co_sitemap_audit_2026-10-13.xlsx",
      "topic": "sitemap_audit",
      "date": "2026-10-13"
    }
  ]
}
```

`run_folder`, optional `template`, and optional `profile` resolve against the manifest's directory. The default template/profile are bundled with the skill. Each source `file` must be a relative path inside the run folder (including resolved symlinks). Only listed files are read. Output goes in `run_folder`. Require exact lowercase hostnames; add www/non-www variants only when the run scope supports both. `asset_hosts` applies to Destination resource URLs only. A report without structured URLs produces an identity-verification warning; a filename alone does not independently establish its host.

`date` is the source evidence/filename date. `output_date` is an explicit report/consolidation date; neither silently overrides the other. Mixed evidence dates are allowed only because the exact reports belong to the explicitly selected run, and remain visible in provenance. No newest-file selection.

Old filenames lacking numeric prefixes need per-source `"legacy_name": true`. Company/topic/date are still checked. A modern prefix must match the actual scope even with legacy permission. A partial source report needs explicit `"checks": ["17.2"]`; otherwise all profile IDs for that topic are required. Multiple disjoint subset reports are supported; overlapping IDs or combined/split candidates stop.

## Verified source families

| Topic key | Expected IDs | Existing combined detail sheet |
|---|---|---|
| robots_audit | 10.1–10.7 | 10. Robot.txt / 10. Robots.txt |
| sitemap_audit | 11.1–11.9 | 11. Sitemap |
| pagination_audit | 14.1–14.6 | 14. Pagination |
| hreflang_audit | 15.1–15.3 | 15. HREFLANG |
| page_titles_audit | 17.1–17.4 | 17. Page Title |
| meta_descriptions_audit | 18.1–18.3 | 18. meta_description |
| headings_audit | 19.1–19.3 | 19. Heading |
| images_audit | 20.1–20.2 | 20. Image |

`https_audit` reserves destination 9.1–9.3 and a provisional source layout; no real HTTPS source was supplied for this version. Do not claim its schema was verified. GSC is deliberately absent. Pagination/Hreflang filenames accept both `onsite_pagination_audit` / `onsite_hreflang_audit` and their shorter registered tokens.

All observed Checklists are header row 1 with `Item No.`, `Item Name`, `Initial Check` or `Check Result`, `Findings`, `Instruction`. Robots instead embeds `10.x` in `Check`, with `Result`, `Findings`, `Instruction`. Headers match registered aliases case-insensitively after trimming/collapsing spaces; column order may vary. Duplicate matching aliases stop.

All observed issue tables have header row 1. Combined tables identify records by `Item No.`, or a leading check ID in `Issue`/`Issues`. Future split sheets begin with a validated child ID (e.g. `17.2 Page titles`); their rows may omit Item No. If it is present, it must agree. A header-only table is valid; an absent table is not equivalent to no issues.

## Destination contract

`scripts/profiles.json` owns exact destination worksheet names. Existing template row positions are discovered from B check IDs, not hardcoded; floating point artifacts such as `11.299999999` normalize only for numeric IDs. Text IDs remain exact (`11.10` never becomes `11.1`). D/E/F must be writable non-formula, non-hyperlink cells. G and populated detail A1 are controlled navigation cells: rebuild their links, replacing stale HYPERLINK formulas only in these cells.

`main_sheet` identifies the input template sheet `SEO Implementation Checklist`; `output_main_sheet` is `(auto) SEO Checklist`, explicitly requested by the user. The full prefixed original title would exceed Excel's 31-character limit. Rename only the generated copy and update internal hyperlinks, formulas, named references and chart series references. Source workbooks and the bundled template retain their original sheet names.

The output status vocabulary is `✔ / ✖ / NA / - / Human check`. Source `Human Check` normalizes only its capitalization. Manual checks retain their reason and unavailable issue count. The status dropdown is extended for onsite cells; series 7 keeps its original dropdown and values. Missing reports use `-` plus `未检查`, not Human check.

Detail headers are row 3, except Robots row 4. Existing notes below headers are preserved; imports start below the last occupied row with a gap. Match existing headings to exact source headings first, then registered semantic aliases. Fill only these fields. Do not append columns, summary blocks or explanatory rows. Original source values, provenance and output row/cell mappings remain in the validation sidecar, including unmapped fields. Stop if an entire detail row maps to no existing field. Main G targets the first populated detail cell for its check, or the existing header for a check with no details. Detail A1 returns to the first corresponding Checklist B row. Exact targets are stored in navigation_links and checked after saving and during verify.

A split destination already exists in the supplied template for each 17/18/19/20 child check. Source naming changes therefore need no destination renaming. Preserve precise native names such as `18. Descriptions - Missing` for 18.1.

## Agent adaptation

For new names/headers, inspect only the explicit files. Copy the profile into the run folder and set `profile` in the manifest. Update source sheet aliases, filename tokens, field aliases or exact targets only when the IDs and column meaning establish the mapping. Never rename one metric to another (for example Character Count to Pixel Width). Extra source fields are preserved in the sidecar, not added to the destination schema or silently interpreted.

Different header rows, non-tabular issue layouts, merged input data, formula-derived results and new status meanings need a tested reader adapter; JSON aliases alone cannot establish their semantics. Unknown result tokens stop. Source formula cells stop; the tool does not execute workbook functions or rely on stale cached results.

Input results remain authoritative. This consolidator does not adjudicate SEO rules or improve findings. Missing reports, manual checks and incomplete evidence remain visible. Counts are observed rows, not a claim that all defects have been found; preserve source units and count limits.

## Current verification limits

The real sample covers eight onsite reports, 37 check results and 1,086 problem-detail rows. GSC is excluded. The template has known inherited broken Summary formulas and Google-specific formulas/anchors. The script checks preservation, requested writes, URL links and workbook navigation; it does not recalculate those original formulas or promise their Excel compatibility.
