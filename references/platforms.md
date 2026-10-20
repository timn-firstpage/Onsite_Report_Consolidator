# Windows and macOS

Use the same skill folder and Python scripts on both systems. Require Python 3.10+ and both packages in `requirements.txt`: openpyxl for Excel and Pillow for preserving embedded template images. Microsoft Excel, MCP, a Google account and Drive credentials are not required for consolidation.

The development machine's interpreter, home directory and drive letters are not portable configuration. Select an interpreter on the current machine. Install dependencies into a venv outside the skill repository, then use that interpreter's absolute path on every invocation. GUI Codex need not inherit terminal activation or environment variables.

## Windows — PowerShell

Substitute the actual skill path. `py -3` must select Python 3.10+; if the launcher is unavailable, use an installed Python executable instead.

```powershell
$skillDir = "C:\path\to\onsite-report-consolidator"
$venvDir = Join-Path $env:USERPROFILE ".venvs\onsite-report-consolidator"
py -3 -m venv "$venvDir"
& "$venvDir\Scripts\python.exe" -m pip install -r "$skillDir\requirements.txt"
& "$venvDir\Scripts\python.exe" "$skillDir\scripts\consolidate.py" doctor
& "$venvDir\Scripts\python.exe" "$skillDir\scripts\consolidate.py" build --manifest "C:\audits\本次 run\run.json"
```

The call operator `&` is necessary when invoking a quoted executable path in PowerShell. Do not treat the Microsoft Store Python alias as proof that a usable interpreter is installed; `doctor` must actually run successfully.

## macOS — zsh / bash

Substitute paths for this Mac; the Windows `C:\Users\...` path cannot be reused.

```bash
skill_dir="/path/to/onsite-report-consolidator"
venv_dir="$HOME/.venvs/onsite-report-consolidator"
python3 -m venv "$venv_dir"
"$venv_dir/bin/python" -m pip install -r "$skill_dir/requirements.txt"
"$venv_dir/bin/python" "$skill_dir/scripts/consolidate.py" doctor
"$venv_dir/bin/python" "$skill_dir/scripts/consolidate.py" build --manifest "/Users/you/audits/本次 run/run.json"
```

Python and Pillow wheels are selected for the current machine architecture. Use an interpreter matching that architecture rather than copying a Windows venv to a Mac or copying an Intel venv to Apple Silicon.

## Shared run manifests and files

- Prefer forward slashes in relative manifest paths, e.g. `"file": "reports/sitemap.xlsx"`, on both platforms. Legacy relative backslashes are normalized; absolute foreign-machine paths are rejected rather than guessed.
- `run_folder`, `template`, and `profile` resolve relative to the manifest. Native absolute paths and `~` home paths are supported for these fields and CLI inputs. Source `file` remains confined to the declared run folder.
- All JSON and CLI output use UTF-8. Filenames with spaces and Chinese characters are supported; quote paths in shell commands. Canonically equivalent NFC/NFD company names are compared equally so macOS filename normalization does not cause false identity conflicts.
- Case-sensitive names remain explicit. A different company, topic, date, ID or candidate report still stops the run. Portability does not relax audit validation.
- The output name, `(auto) SEO Checklist`, Human check results, GSC exclusion and internal/website hyperlink behavior are identical across platforms.
- Windows publishes the validated temporary workbook by a non-replacing rename. macOS/POSIX normally uses a non-replacing hard link. On volumes that do not support hard links, an exclusive-copy fallback prevents overwrite and removes its own partial file on an I/O error. This fallback is not atomic; after an interrupted process, do not treat an output without a matching valid sidecar as verified. Resolve it explicitly rather than overwrite it automatically.
- `verify` checks the saved workbook against its matching `.validation.json`. This is file validation, not native Excel UI/recalculation validation.

## Automated tests

`.github/workflows/cross-platform.yml` runs the same suite on Windows with Python 3.10 and 3.12, Intel macOS (`macos-15-intel`) with Python 3.10, and Apple Silicon macOS (`macos-latest`) with Python 3.12. The suite includes a CLI build/verify from an unrelated working directory, Chinese/space paths, a preserved embedded image, manual-check statuses, path confinement, Unicode normalization and publication failures. Native foreign-path tests intentionally skip on the other OS. A workflow definition alone is not a passing run; inspect its actual Actions result before claiming macOS validation.

The tests contain synthetic input only. No service account, customer source report, browser login or Google API is involved.
