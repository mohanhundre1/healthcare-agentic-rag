# DEF Natural Gas — formula output update

The first flow reads the seed workbook, validates and calculates records through the existing common-library integration, then creates `I_FPC_OA.xlsx`. Its `I_FPC OA` sheet contains the 826 formulas captured from the supplied client `Actual_data(1).xlsm` rather than only materialized results.

**Formula references are intentionally internalized.** For example, `='[17]GSNG Imbalance'!Y11` becomes `='GSNG Imbalance'!Y11`. The expression, destination address and referenced source address are otherwise unchanged. This is not byte-identical preservation of obsolete external workbook links. The original expressions remain in `def_fpc_oa/client_formulas.py` for review.

Three visible dependency sheets contain validated source snapshots: `FGT Data Download - New`, `TOTAL FGT OA`, and `GSNG Imbalance`. They contain the inputs needed by this first flow, not the complete upstream client workflows. No template is required at runtime. Downstream journal and allocation work is outside this update.

## Setup and run on Windows

Use the approved Python 3.13+ environment and private Artifactory configuration for the supplied common package. This change does not add Python dependencies or modify the common library.

```powershell
python -m pip install -r requirements.txt
$env:FPC_OA_RECALC_ENGINE = "excel"
python run_local.py --seed "input\IMBALANCES_Seed_Data.xlsx" --output-dir output_formulas
```

The Excel path requires installed, licensed desktop Excel and PowerShell. `def_fpc_oa/recalculate_excel.ps1` opens the temporary generated workbook with macros disabled, calculates, saves and closes it. Organization script policy must permit the approved script. Configure an approved calculation environment for deployed agents; a headless server without Excel needs a different supported engine.

The supported alternative is an installed LibreOffice Calc executable:

```powershell
$env:FPC_OA_RECALC_ENGINE = "libreoffice"
$env:FPC_OA_SOFFICE = "C:\Program Files\LibreOffice\program\soffice.exe"
python run_local.py --seed "input\IMBALANCES_Seed_Data.xlsx" --output-dir output_formulas_calc
```

Default `auto` uses Excel on Windows and LibreOffice elsewhere. No engine is installed automatically. Formula output fails instead of publishing missing or mismatched calculated results. Existing output files are not overwritten.

## What the writer checks

1. Validated raw inputs reproduce the supplied Python result using strict calculations.
2. The native engine calculates every target formula.
3. All 826 saved target formula results match Python (including dates, title, blanks and NA).
4. Original formula text is restored with only the documented external-prefix substitution, retaining genuine engine-calculated caches. This prevents an engine's syntax normalization from changing the client formula expressions.
5. Formula text is read back and checked before atomic output publication.

Python calculations remain the independent check and JSON handoff. Excel cells keep numeric signs, bracket display formats, zero values, N/A presentation and blank zero-burn BTU factors. Client B/AA formulas return date text using `TEXT(...,"mm/dd/yyyy")`, matching the client's formula behavior.

Raw dependency records start at row 4, matching the client's `$A$4:$R$1048576` lookup range. The supplied seed begins at row 3: all validated records, including the first, are copied into the generated dependency sheet from row 4. No raw record is dropped.

## Files changed

- `def_fpc_oa/output.py`: layout, formula writing, validated-input requirement and verification.
- `def_fpc_oa/client_layout.py`: captured client labels and cell styles.
- `def_fpc_oa/client_formulas.py`: original 826 formulas and explicit link localization.
- `def_fpc_oa/formula_workbook.py`: dependency sheets, native calculation and cache/value validation.
- `def_fpc_oa/recalculate_excel.ps1`: Windows Excel calculation adapter.
- `def_fpc_oa/inputs.py`: retain the actual source period label for the title formula.
- `def_fpc_oa/workflow.py`, `run_local.py`: pass validated raw inputs into the writer.
- `tests/test_formula_output.py`: native formula and handoff checks.
- `verify_formula_output.py`: diagnostic harness without the server/runtime.

Custom callers must now pass the raw validated payload:

```python
write_output(output_path, calculated_result, validated_input=payload)
```

## Checks available here

50 tests: 47 passed and 3 common-runtime tests skipped because the full shared runtime is unavailable. Native tests used LibreOffice Calc. They cover every formula expression, result readback, first raw-record inclusion, edited inputs, Excel half-up rounding, signed burns, blank BTUs and next-stage summary parity. Invalid engine configuration and mismatched inputs do not publish a workbook.

The supplied August workbook passed 774 saved-seed comparisons and all 826 target formula-result comparisons. The real common workbook-reader source module was loaded directly for this diagnostic; the package/framework was not mocked. The next-stage reader accepted 31 records and produced the same I_Imbalance summary.

`output/I_FPC_OA.xlsx` is the tested formula example. Its report states that these are diagnostic checks, not RequestPipeline or HTTP A2A acceptance. Desktop Excel/PowerShell acceptance must be run in the approved Windows environment. Existing calculations and downstream reader code are unchanged.

```powershell
python -m unittest discover -s tests -v
```

A diagnostic run using an installed common reader:

```powershell
python verify_formula_output.py --seed "input\IMBALANCES_Seed_Data.xlsx" --output-dir diagnostic_formulas --engine excel
```

Only when the complete common package cannot import, `--common-root <path-to-common-source>` explicitly loads its actual workbook-reader module. That option is diagnostic and does not validate the common RequestPipeline.

## Scope limits

The formula snapshot and daily writer support this supplied 31-day layout. Formulas that cannot evaluate, including a client `AVERAGE` over entirely blank factors, are reported rather than silently replaced by new formulas. Any input edit requires native recalculation and saving before downstream data-only readers see refreshed results. The three embedded dependency sheets are source snapshots; they do not regenerate the upstream GSNG/TOTAL workflows. No macros execute, and this update does not refresh or validate the full client journal workbook.
