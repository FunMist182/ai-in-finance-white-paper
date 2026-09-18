# AI in Finance: White Paper

Research materials accompanying [AI in Finance: White Paper](AI_in_Finance_White_Paper.pdf). The PDF is the main paper; the [Word version](AI_in_Finance_White_Paper.docx) is retained for editing.

## Case studies

- **Case Study 1 — Tesco and Carillion accounting-risk screening.** Fixed Python rules examine accruals, cash conversion and working-capital growth. The Tesco screen flags some earlier movements but does not establish advance detection of the commercial-income issue. Carillion shows recurring receivables warnings and periods of negative operating cash flow despite reported profits. These are retrospective investigation prompts, not trained machine-learning predictions.
- **Case Study 2A — LSEG valuation.** A Python-generated workbook contains a five-year DCF, comparable-company screen, scenarios and sensitivities. The paper reports conditional downside/base/upside DCF values of £95.41/£122.92/£150.91 per share. They depend on the model's assumptions and historical reference inputs, not current market conditions.
- **Case Study 2B — LSEG model governance.** The paper examines calculation checks, source definitions, assumptions and unresolved limitations of the same model. Passing arithmetic checks does not establish that assumptions are economically appropriate or that the ownership bridge is complete.

## Reproducibility status

**This collection currently contains the supplied original scripts, not the complete newer Python bundle cited in the PDF.** The paper refers to `AI_in_Finance_Python_Work.zip` and an extracted `Python_work` folder. That bundle was not among the available files when this repository was assembled.

Missing items include `tesco_analysis.py`, `carillion_analysis.py`, `lseg_case_outputs.py`, `validate_lseg.py`, the later Tesco deterioration output, and the original LSEG independent-review, spreadsheet-recalculation and scenario-switch records. Do not treat the included files as evidence that every check described in the paper has been reproduced. The historical review records cannot be recreated retrospectively as if they were originals.

Repository preparation checks successfully ran all three supplied scripts in a separate temporary copy. Both accounting CSVs matched the supplied outputs, and all 277 saved formula values from the rebuilt LSEG workbook matched the supplied workbook. These are reproduction checks, not independent model validation or spreadsheet-engine recalculation. Details and package versions are in `results/packaging_checks.json`.

One concrete version difference: the original `Tesco code.py` stores early unavailable accrual-history comparisons as `False`; the newer implementation described in the paper correctly treats them as unavailable. Those early values must not be interpreted as passed tests. The original files are retained unchanged for traceability.

## Files and folders

| Location | Contents |
| --- | --- |
| `AI_in_Finance_White_Paper.pdf` | Main paper |
| `AI_in_Finance_White_Paper.docx` | Editable paper |
| `Case study 1/Tesco code/` | Individual original Tesco and Carillion Python scripts |
| `Case study 1/Tesco_Master_Dataset_2010_2015.xlsx` | Required Tesco input data |
| `Case study 1/*.csv` and `*.png` | Supplied accounting-screen outputs and Tesco charts |
| `Case study 2/Davids version/` | LSEG Python builder and generated valuation workbook |
| `results/source_manifest.json` | SHA-256 fingerprints of unchanged supplied files |
| `results/packaging_checks.json` | Checks performed specifically for repository preparation, when present |
| `requirements.txt` | Python dependencies and tested versions |

The original folder layout is preserved because the Tesco script finds its workbook one folder above the script and the LSEG builder saves its workbook beside its source. There is no duplicate `code/` or `data/` tree.

## Run the supplied code

Use Python 3.12. From the repository root, install the dependencies:

```sh
python -m pip install -r requirements.txt
```

Run the accounting screens:

```sh
python "Case study 1/Tesco code/Tesco code.py"
python "Case study 1/Tesco code/Carillion code.py"
```

The scripts display charts; close each chart window to continue. For non-interactive execution, set the `MPLBACKEND` environment variable to `Agg` before running them. In PowerShell: `$env:MPLBACKEND = "Agg"`.

Tesco writes its CSV and two charts in `Case study 1/`. Carillion embeds its input data and writes its CSV and two charts in `Case study 1/Tesco code/`; the supplied Carillion CSV one folder above is a retained reference copy. Re-running scripts overwrites their generated files.

Rebuild the LSEG workbook:

```sh
python "Case study 2/Davids version/build_lseg_valuation_python_only.py"
```

This overwrites the adjacent workbook. The model's inputs and source register are embedded in the builder/workbook; no live data service or API key is required. Python supplies saved formula values. Rebuilding those values is not an independent spreadsheet-engine recalculation or an independent implementation of the model.

## Interpretation and limitations

Accounting amounts are GBP millions. Columns ending `_YoY_%` store fractional changes: `0.05` means 5%. A missing first-year accrual ratio reflects the absence of prior-year assets. Cash conversion can be misleading with very small or negative profits. Carillion's script records its mixed reporting vintages and source limitations.

The LSEG exercise uses FY2025 financials and a 26 February 2026 reference date; the programme-price and voting-share proxies were disclosed on 27 February, so this is retrospective. The base DCF relies heavily on terminal value, peer inputs remain partly approximate, and the equity bridge is simplified. See the paper and workbook source register for the full limitations. This is a research and educational project, not a current investment recommendation.

Older investment/transaction narratives, personal presentation guidance, backups and editorial-review material are excluded. The transaction narrative is not the governance case presented as Case Study 2B in the final PDF. Third-party annual-report PDFs are not duplicated here; consult the source references in the paper, scripts and workbook.

No reuse licence has been assigned in this repository.
