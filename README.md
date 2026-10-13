# LDI / ALM Lab

Rebuilding the core of the SOA *Liability-Driven Investment Benchmark Model*
(Shang & Hossen, 2019) in Python, checking every step against the original Excel tool.

## Status
Work in progress. Phase 1: liability side.

## How to run
```bash
pip install -r requirements.txt
python scripts/export_from_excel.py path/to/liability-driven-investment.xlsm
pytest
```

## Structure
- `scripts/` one-off tools (export inputs from the Excel model)
- `src/ldi/` the model code
- `tests/` checks that Python matches Excel
- `data/` inputs exported from the workbook
- `notebooks/` exploration and charts
