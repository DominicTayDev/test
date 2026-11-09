# Singapore card interchange++ workbook

`SG_Card_Interchange_Rates.xlsx` maps Singapore-issued cards (bank, card, scheme, tier) to Visa / Mastercard
interchange groups and computes interchange++ (interchange + scheme fee + markup) once rates are entered.

- **Rate Table**: yellow cells are inputs; rates are not public for Singapore and are intentionally left blank.
- **Card Catalog**: 201 cards from public sources (Nov 2026), each with source URL and confidence.
- **Merchant Pricing**: blended IC++ cost for a merchant's card mix vs a quoted flat MDR.
- **Card Data Capture**: per-transaction fields needed to know each card's actual interchange.

Rebuild from the source CSVs:

```bash
python3 tools/build_xlsx.py data SG_Card_Interchange_Rates.xlsx
```

Note: rebuilding overwrites any rates typed into the workbook.
