# Verification — 2026-10-15

Windows, bundled Python 3.12 / openpyxl. Current runtime dependencies are Python 3.10+, openpyxl 3.1.5+ and Pillow 10.4+ (see requirements.txt for supported upper bounds). The development repository is separate from installed skills.

- Offline suite: 42 tests, including combined and future split layouts; raw status fidelity with incomplete evidence; filename/company/date/check coverage; duplicate selection and conflicting split IDs; wrong page hosts versus allowed resource destinations; required/missing/unnamed fields; merged inputs; source formulas; source count reconciliation; literal URL validation; template formulas/styles/static notes; unchanged GSC; missing checks; navigation; website links; provenance; output collision and later tampering. Added explicit Human check/dropdown distinction, main-sheet rename/reference and active-sheet persistence tests, including the provided template's absent workbook view.
- Real provided historical ZIP: eight explicitly selected onsite reports, 37 check results, 1,086 problem-detail rows. The GSC workbook was excluded completely. Missing template checks use `-` / `未检查`, except GSC and the non-status CMS input. Only declared local files were read.
- Output verification: 197 native internal hyperlinks and 2,134 imported website/resource links. All written cell values and hyperlinks are checked after reopening; untouched effective cell styles, values, formulas, links, comments, merges, data validation, sheet order and image/chart counts are compared to the template after the explicitly authorized main-sheet rename and onsite dropdown extension. The source template remains unchanged.
- Independent review/forward test: inspect → build → verify succeeded on a native-layout fixture and explicit subset; a repeated build refused overwrite. Three reader defects were independently reproduced (unnamed evidence cells, plural count claims, merged inputs); regression tests failed before the fixes and the full suite passed afterward.
- Template roundtrip normalization: duplicate style indexes and equivalent boolean defaults are compared by effective style. Numeric template serialization noise is normalized to 14 significant digits for preservation comparisons; source IDs, result tokens and integer evidence counts are checked exactly.
- Official skill `quick_validate.py`: valid. PyYAML was installed only in an ignored validation directory, not added to runtime dependencies.
- Thirty-three preview ranges were rendered across the main Checklist and all 16 populated detail worksheets, including summary blocks after long evidence tables. Relevant main/detail/summary previews were inspected; row heights wrap new content. Imported evidence tables remain in their native template columns. Template screenshots are explicitly labelled reference-only.

## Limits

The supplied template already has 69 formula/anchor diagnostics: Summary missing-sheet/#REF references, Google-export placeholder formulas and Google `heading`/`gid` anchors. They remain unchanged and are listed in the run's `.validation.json`; the report does not claim those original features work in native Excel. The preview renderer emits the inherited missing-sheet diagnostics and returns a nonzero process status despite producing all requested PNGs; it is not used as formula/recalculation proof.

At the October 15 baseline, native Excel recalculation/click testing, actual macOS execution, HTTPS source-schema integration and Drive input/download/upload were not performed. No service account file or credential was read. Customer source reports, generated reports, manifests and previews are excluded from Git.

## Cross-platform update — 2026-10-20

- A fresh external Windows venv (Python 3.12.14, openpyxl 3.1.5, Pillow 13.0.0) passed 54 tests: 53 passed and one native macOS path test intentionally skipped. Official skill validation passed.
- A real historical-run regression preserved all four template images, 37 check results and 1,086 detail rows. Reopening and verification passed with 197 internal and 2,134 website/resource hyperlinks. The 69 inherited template diagnostics remain reported.
- Hosted CI for commit `b66bd7e` passed on Windows Python 3.10, Windows Python 3.12, and Apple Silicon macOS Python 3.12. Intel macOS Python 3.10 was still queued when this record was written. Consult the [actual run](https://github.com/timn-firstpage/Onsite_Report_Consolidator/actions/runs/39357223847) for its current status.
- Tests cover fresh installation/image support, a read-only doctor, native and foreign absolute paths, relative separators, Chinese/spaced paths, NFC/NFD names, CLI invocation outside the repository and no-overwrite output publication/fallback cleanup. Missing Pillow now stops before template images can be silently dropped.
- Native Excel UI/recalculation and Drive operations remain untested. This update does not read credentials or alter the uploader skill.
