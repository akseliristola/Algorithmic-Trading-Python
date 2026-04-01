# AGENTS.md

## Cursor Cloud specific instructions

This is a Python-based S&P 500 Stock Analysis Toolkit that computes equal-weight portfolios, momentum rankings, and value rankings, outputting results to Excel files.

### Project structure

- `src/main.py` — Entry point; calls all three strategies sequentially.
- `src/constants.py` — Loads `sp_500_stocks.csv` and defines shared constants (`stocks`, `chunks`, `portfolio_size`).
- `src/SP500EqualWeight.py` — Equal-weight S&P 500 portfolio calculator.
- `src/QuantitativeMomentumStrategy.py` — Momentum strategy (scipy percentile scoring).
- `src/quantitativeValueStrategy.py` — Value strategy (P/E, P/B, P/S, EV/EBITDA, EV/GP).
- `src/secretss.py` — **Gitignored.** Must exist with `IEX_CLOUD_API_TOKEN = '<token>'` for the app to import. Create manually.

### Running the application

Run from the `src/` directory: `cd src && python3 main.py`. Output goes to `../excel/`.

### Key caveats

- **IEX Cloud API is defunct.** The external API at `cloud.iexapis.com` is no longer operational. All three strategy modules call this API, so `main.py` will fail at runtime with HTTP errors. To test core computation logic (pandas, scipy percentile scoring, xlsxwriter output), use mock data instead of live API calls.
- **`scipy` is an implicit dependency** used by `QuantitativeMomentumStrategy.py` and `quantitativeValueStrategy.py` but not listed in `requirements.txt`. It is installed alongside `requirements.txt` in the update script.
- The `excel/` output directory must exist before running. It is not created automatically by the scripts.
- All source files run from `src/` and use relative paths (`../sp_500_stocks.csv`, `../excel/`).
- No linter, test framework, or CI is configured in this project. Use `python3 -m py_compile <file>` for syntax checks.
