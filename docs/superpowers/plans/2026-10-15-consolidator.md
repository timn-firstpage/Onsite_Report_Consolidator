# Onsite Report Consolidator Implementation Plan

> **For agentic workers:** use executing-plans to implement this plan inline; user approved building after the mapping conversation.

**Goal:** Independently consolidate explicitly selected onsite Excel reports into the preserved template.

**Architecture:** Manifest selects the run; profiles map topics/IDs/columns; input validation produces normalized check/evidence records; writer fills existing sheets and verifies a temporary output before publication.

**Tech Stack:** Python 3.10+, openpyxl 3.1.5+, unittest; Windows/macOS portable paths.

**Spec:** ../specs/2026-10-15-consolidator-design.md

## Global Constraints
- Exclude series 7/GSC; no Drive or credential operations.
- Status ✔/✖/NA/-; preserve source conclusions and Findings.
- Output Autocheck_onsite_audit_{companyname}_{date}.xlsx beside run inputs; refuse overwrite.
- Template content outside declared writes remains intact; unknown mappings stop.

## Review Focus
- Old filename compatibility must be explicit, not a universal filename bypass.
- Split-sheet IDs and embedded IDs must agree; duplicates must stop.
- Formula cells and hyperlinks outside designated data/navigation writes survive.
- A template instruction/example is not an issue record.
- Incomplete evidence must not rewrite explicit results or imply a verified zero.

### Task 1 — normalized inputs and profiles
Files: scripts/inputs.py, scripts/profiles.json, tests/test_consolidator.py.
Interface: read_run(manifest_path) returns run metadata, checks and evidence grouped by check.
- [x] Write fixture tests for valid, missing, duplicate, split, schema, count, identity and filename failures; run RED.
- [x] Implement exact manifest selection and explicit profile aliases; run GREEN.

### Task 2 — template writing and validation
Files: scripts/workbook.py, scripts/consolidate.py; same tests.
Interface: build(manifest_path) returns output path and verification report; CLI inspect/extract/build/verify.
- [x] Test real template preservation, exact ID routing, statuses/counts, navigable URLs, collision and tamper detection; run RED.
- [x] Implement controlled writes, traceability, verification and exclusive publication; run GREEN.

### Task 3 — reusable skill and real-data integration
Files: SKILL.md, references/mapping.md, README.md, assets/template.xlsx.
- [x] Document profile adaptation and validated commands, excluding GSC and Drive.
- [x] Run supplied ZIP through build and verify; inspect sample output and render every populated sheet.
- [x] Run full offline suite, skill format validation and independent forward evaluation; record limitations.
