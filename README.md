# cpm

Cashflow / CPM project assets and automation.

## Weekly status deck

Source data and generator live under `weekly/`:

| File | Purpose |
|------|---------|
| `CPM-ProcurementForecasting-Status.xlsx` | Backlog export (sheet `Backlog`) |
| `generate_cashflow_status_pptx.py` | Builds the status PowerPoint |
| `assets/aes-logo.png` | AES logo for title and closing slides |
| `CASHFLOW_Status_YYYYMMDD.pptx` | Generated decks |

### Regenerate

```bash
cd weekly
python3 generate_cashflow_status_pptx.py
```

Edit the report date in `main()` inside the script when needed.

Requires: `openpyxl`, `python-pptx`, `Pillow`.

```bash
pip install openpyxl python-pptx Pillow
```
