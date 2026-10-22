# DEF FPC OA — formula-only output

This version generates `I_FPC_OA.xlsx` containing the 826 client formulas and client layout, without running Excel, LibreOffice, PowerShell or any formula calculation engine. `recalculate_excel.ps1` has been removed.

The old `[17]` external workbook prefixes are replaced with internal sheet references. The expressions and destination addresses are otherwise preserved. Three visible dependency sheets contain the required validated source snapshots. No runtime template is required.

## Run

Use the existing approved Python 3.13+ virtual environment and common-package/private Artifactory configuration. Requirements are unchanged.

```powershell
python run_local.py --seed "input\IMBALANCES_Seed_Data.xlsx" --output-dir output_formulas_only
```

No `FPC_OA_RECALC_ENGINE` variable or desktop spreadsheet engine is required. Existing engine environment variables have no effect in this version.

## Continue to I_Imbalance

Formula cells have no saved results. The XLSX cannot be used with the existing data-only next-stage reader until someone calculates and saves it. Use the calculated JSON instead:

```powershell
python run_i_imbalance.py --input "output_formulas_only\calculated_I_FPC_OA.json" --output-dir output_imbalance_from_json
```

Python calculations still supply the signed business values in JSON. No formula-result caches are fabricated from those values. Excel's formula cells may look blank or show unavailable results until calculated in a spreadsheet application. This is an expected consequence of formula-only generation.

The generated workbook requests manual calculation, disables full calculation on load and disables calculation on save. Excel's application-wide settings or subsequent user actions can override workbook preferences; this does not enforce application behavior.

## Files

- `def_fpc_oa/output.py`: writes layout, inputs and formulas; requires `validated_input=payload`.
- `def_fpc_oa/client_layout.py`: client labels and display formats.
- `def_fpc_oa/client_formulas.py`: original client formulas and explicit external-prefix replacement.
- `def_fpc_oa/formula_workbook.py`: embedded dependencies and formula-text verification; no recalculation code.
- `def_fpc_oa/inputs.py`: retains source period label.
- `def_fpc_oa/workflow.py`, `run_local.py`: pass validated inputs to the writer.
- `tests/test_formula_output.py`: formula-only output, absence of caches, JSON handoff and input consistency checks.
- `verify_formula_output.py`: diagnostic harness; no server required.

The Python calculation rules, common library and I_Imbalance calculation/reader code are unchanged. The supplied layout is limited to 31 daily rows. The embedded dependency sheets are snapshots, not regenerated upstream workflows.

## Verification

49 tests: 46 passed, 3 common-runtime tests skipped because the full shared runtime is unavailable here. Checks assert that no external calculation command runs, all 826 formula strings match the localized client snapshot, formula results remain uncached, and calculated JSON produces the same I_Imbalance summary. The supplied August diagnostic also passed 774 saved-seed comparisons and the 31-record JSON handoff.

```powershell
python -m unittest discover -s tests -v
python verify_formula_output.py --seed "input\IMBALANCES_Seed_Data.xlsx" --output-dir diagnostic_formula_only
```

Only when the full common package cannot import, `--common-root <common-source-root>` explicitly loads the actual shared workbook-reader module. The supplied example under `output/` was produced with that diagnostic route, not a simulated RequestPipeline. Live A2A and RequestPipeline acceptance were not run here. No Excel-formula result validation is claimed for this formula-only output.
