# Purchase Order Information Extractor

An Excel-based automation solution that extracts structured Purchase Order (PO) information from PDF documents and generates formatted Excel reports with a single click.

The application combines  to automate manual data extraction and reporting workflows while maintaining a centralized dashboard for generated Purchase Orders.
---

## Overview

Purchase Order Information Extractor is designed to simplify the process of converting Purchase Order PDFs into structured Excel reports.

Instead of manually copying PO details from PDF documents, users can select a PDF from the dashboard and generate a fully formatted Purchase Order sheet automatically. The application also maintains a history of generated reports for easy tracking.

---
https://github.com/user-attachments/assets/b3e6c580-877f-423f-88ea-6a4887729104

## Key Features

- Automated Purchase Order data extraction from PDF files
- One-click Excel report generation through VBA
- Dashboard-driven workflow
- Automatic worksheet creation for each Purchase Order
- Purchase Order history management
- Grand Total calculation
- Automatic formatting and styling
- Auto Filter and Freeze Panes
- Dynamic column width adjustment
- Excel-native workflow using xlwings

---

## Tech Stack

| Technology | Purpose |
|------------|---------|
| Python | Core application logic |
| pdfplumber | PDF text extraction |
| openpyxl | Workbook formatting |
| Regular Expressions | Data extraction and parsing |
| Microsoft Excel | User Interface |

---

## Project Architecture

```
                                                              Purchase Order PDF
                                                                      │
                                                                      ▼
                                                             PDF Text Extraction
                                                                      │
                                                                      ▼
                                                             Purchase Order Parser
                                                               (Regex + Python)
                                                                      │
                                                     ┌───────── ──────┴────────────────┐
                                                     ▼                                 ▼
                                              Purchase Order Details            Material Details
                                                     │                                 │
                                                     └────────────────┬────────────────┘
                                                                      ▼
                                                             Excel Report Generator
                                                          (Visual Basic for Application)
                                                                      │
                                                                      ▼
                                                         Formatted Purchase Order Sheet
                                                                      │
                                                                      ▼
                                                           Dashboard History Update
```



## Application Workflow

1. Launch the Excel Dashboard.
2. Browse and select a Purchase Order PDF.
3. Click **Generate Sheet**.
4. Python extracts Purchase Order information from the selected document.
5. A new worksheet is created automatically.
6. Purchase Order and Material information are written into the report.
7. Dashboard history is updated.
8. Report generation status is displayed.

---

## Extracted Information

### Purchase Order Details

- Purchase Order Number
- Purchase Order Date
- Vendor Name
- Vendor Code
- Currency

### Material Information

- Material Number
- Material Description
- Order Quantity
- Unit
- Price
- Net Value

---

## Dashboard

The dashboard serves as the primary interface for the application.

It provides:

- PDF selection
- One-click report generation
- Generation status
- Purchase Order history
- Sheet tracking

---

## Report Features

Each generated worksheet includes:

- Purchase Order summary
- Material information table
- Automatic Grand Total
- Professional formatting
- Auto Filter
- Freeze Panes
- Dynamic column sizing
- Cell alignment and borders

---


## Requirements

- Windows 10 / Windows 11
- Microsoft Excel Desktop Version
- Python 3.10+


---

The application automatically creates a new worksheet and updates the dashboard.

---

## Future Enhancements

- Batch PDF processing
- OCR support for scanned PDFs
- Database integration
- CSV export
- Email report generation
- Logging and diagnostics
- User authentication

---
## Repository Notice

This repository contains the documentation for the Purchase Order Information Extractor project.

The complete implementation is maintained in a private repository as it contains automation workflows developed for an organizational use case.

The public repository is intended to showcase the project architecture, workflow, technologies used, and key features while respecting project confidentiality.
