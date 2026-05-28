# AGENTS.md — cpm

## Weekly Cashflow status

- Source: `weekly/CPM-ProcurementForecasting-Status.xlsx` (sheet `Backlog`)
- Generator: `weekly/generate_cashflow_status_pptx.py`
- Output: `weekly/CASHFLOW_Status_YYYYMMDD.pptx`
- Remote: https://github.com/solzito148/cpm

After any weekly status update, **commit and push to `main`** on `solzito148/cpm` (see `.cursor/rules/cpm-weekly-status.mdc`).

```bash
cd weekly && python3 generate_cashflow_status_pptx.py
```

Dependencies: `openpyxl`, `python-pptx`, `Pillow`.
