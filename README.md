# Patent Portfolio Processing Pipeline

Python-based data processing pipeline for transforming raw patent datasets into standardized Excel workbooks with automated formatting, normalization, and portfolio analytics.

## Features

- Supports Excel, CSV, and JSON patent datasets
- Configurable column mapping and normalization
- Automated Excel formatting and styling
- Patent hyperlink generation
- Jurisdiction classification (US / Non-US)
- Portfolio overview and analytics generation
- INPADOC family analysis
- Pivot-style reporting tables
- Conditional formatting and date normalization
- Structured logging and CLI workflow

## Tech Stack

- Python
- pandas
- openpyxl
- watchdog
- JSON configuration

## Architecture

```text
Input Dataset
    ↓
File Loader
    ↓
Column Mapping & Normalization
    ↓
Data Enrichment
    ↓
Excel Formatting Pipeline
    ↓
Overview & Analytics Generation
    ↓
Final Workbook Export
```

## Project Structure

```text
MLConverter/
│
├── config/
│   └── column_rules.json
│
├── modules/
│   ├── column_mapper.py
│   ├── file_loader.py
│   ├── formatter.py
│   ├── output_handler.py
│   └── overview/
│       └── overview_handler.py
│
├── main.py
├── requirements.txt
└── README.md
```

## Notes

Designed for patent portfolio analysis workflows involving large structured datasets, automated Excel reporting, and portfolio-level analytics generation.

