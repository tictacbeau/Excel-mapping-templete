# Remittance Allocation Engine

**Offline, Excel-connected desktop application for automating AR remittance allocation breakdowns.**

Replaces manual copy → paste → transpose → VLOOKUP workflows with a template-driven, wizard-configured system that handles any payor format automatically.

---

## Overview

Each payor sends payment remittance data in their own consistent format. This tool maps those formats to your internal invoice data and produces presentation-ready allocation breakdown sheets — automatically.

```
Payor Remittance File  →  File Converter  →  Normalized Data
    (xls/xlsx/csv/html-as-xls)                  (clean xlsx)
                                                       ↓
                                            Template Engine  →  Completed Allocation Breakdown
                                                       ↑
                                           Internal Data Model
                                         (Power Query, 300K+ records)
```

---

## Features

### Input File Support
| Format | Handling |
|--------|----------|
| `.xlsx` | Read directly |
| `.xls` (genuine OLE2) | Convert via Excel COM |
| `.xls` (fake HTML) | Detected via magic bytes; HTML table parsed and converted |
| `.csv` | Auto-delimiter detection (comma, tab, pipe, semicolon); quoted fields handled |

### Template System
- **Two templates per payor**: Single-Invoice and Multi-Invoice
- **Dynamic body expansion**: Multi-invoice template clones body section N times for N invoices, preserving all formatting
- **Full formatting control**: Font, fill, borders, number format, alignment, column widths, row heights, conditional formatting, print settings, freeze panes

### Payor Setup Wizard (5 Steps)
1. **Identity** — name, aliases, expected file type
2. **Format Detection** — invoice layout (single-cell or one-per-row), regex pattern extraction with live test button
3. **Template Builder** — color-coded cell mapper (DATA/LOOKUP/CALC/BLANK/LABEL), formatting panel, live preview
4. **Data Linkage** — map internal Data Model fields to template cells, configure match rules
5. **Review & Save** — side-by-side preview and save

### Internal Data Model
- Stores 300K+ internal invoice records using Excel's **Data Model** (Power Pivot compressed columnar storage)
- **Refresh wizard** imports new data from any `.xlsx` or `.csv` source
- Schema drift detection (reports added/removed columns)
- Automatic backup + rollback capability
- In-memory Dictionary cache for fast per-invoice lookups

### Processing Workflow (Runtime)
1. Select payor
2. Browse for remittance file
3. Engine auto-converts file format if needed
4. Auto-detects single vs. multi-invoice
5. Expands template, populates data, runs lookups
6. Applies all formatting, formulas, print settings
7. Flags lookup failures (red), duplicate invoices (orange)
8. Outputs completed allocation worksheet

---

## Repository Structure

```
Excel-mapping-templete/
├── src/
│   ├── vba/
│   │   ├── modules/              # VBA standard modules
│   │   │   ├── modConstants.bas  # All app constants
│   │   │   ├── modUtils.bas      # JSON parser, regex, file helpers, audit log
│   │   │   ├── modFileConverter.bas  # Input file detection & conversion
│   │   │   ├── modPayorRegistry.bas  # Payor CRUD + template storage
│   │   │   ├── modTemplateEngine.bas # Template expansion & population
│   │   │   ├── modDataLookup.bas     # Data Model / Power Query integration
│   │   │   ├── modMain.bas           # Entry points, dashboard, orchestration
│   │   │   ├── modRibbon.bas         # Custom ribbon button callbacks
│   │   │   └── modThisWorkbook.bas   # Workbook event handler stubs
│   │   │
│   │   ├── classes/              # VBA class modules
│   │   │   ├── clsCellFormat.cls # Cell formatting properties object
│   │   │   ├── clsPayor.cls      # Payor configuration object
│   │   │   ├── clsRemittance.cls # Parsed remittance data object
│   │   │   └── clsTemplate.cls   # Template layout definition object
│   │   │
│   │   └── forms/                # VBA UserForms
│   │       ├── frmProcessRemittance.frm  # Process remittance dialog
│   │       ├── frmPayorWizard.frm        # New/edit payor wizard (5 steps)
│   │       ├── frmTemplateBuilder.frm    # Interactive template builder
│   │       ├── frmPayorManager.frm       # Payor list + CRUD
│   │       └── frmDataRefresh.frm        # Data Model refresh wizard
│   │
│   └── customui/
│       └── customUI14.xml        # Excel ribbon definition (Office 2010+)
│
├── build/
│   ├── build.ps1                 # PowerShell build script (Windows, requires Excel)
│   ├── build.py                  # Python validation/report script (cross-platform)
│   └── requirements.txt          # Python optional dependencies
│
├── dist/                         # Built .xlsm output (gitignored)
│
└── README.md
```

---

## Building the XLSM

### Option A: PowerShell (Recommended — Windows + Excel required)

```powershell
cd build
.\build.ps1
```

This will:
1. Launch Excel via COM automation
2. Import all VBA modules, classes, and forms
3. Configure `Workbook_Open` event handlers
4. Save as `dist/RemittanceAllocationEngine.xlsm`
5. Inject the custom ribbon XML

Output: `dist/RemittanceAllocationEngine.xlsm`

```powershell
# Specify custom output path
.\build.ps1 -OutputPath "C:\Tools\RemittanceEngine.xlsm"

# Skip ribbon injection
.\build.ps1 -SkipRibbon

# Verbose output
.\build.ps1 -Verbose
```

### Option B: Manual import (any Windows machine with Excel)

1. Create a new `.xlsm` workbook in Excel
2. Press `Alt+F11` to open the VBA IDE
3. Right-click the project → **Import File**
4. Import in this order:
   - All `.bas` files from `src/vba/modules/`
   - All `.cls` files from `src/vba/classes/`
   - All `.frm` files from `src/vba/forms/`
5. In `ThisWorkbook` module, add event handlers from `modThisWorkbook.bas` comments
6. Save as `.xlsm`

### Option C: Python validation (no Excel required)

```bash
cd build
python build.py              # Run all checks
python build.py --validate   # VBA syntax validation only
python build.py --stats      # Lines of code count
python build.py --check-files # Verify all files present
```

---

## First Run Setup

After opening the `.xlsm` for the first time:

1. **Enable macros** when prompted
2. The Dashboard sheet opens automatically
3. Click **Internal Data → Refresh Data** to import your billing system export
4. Click **Payors → New Payor** to configure your first payor
5. Click **Process Remittance** to process your first file

---

## Payor Configuration Reference

### Invoice Layout Types

**`ONE_PER_ROW`** (default): One invoice per row in a column
```
Row 5: INV-1234567   $5,000.00
Row 6: INV-1234568   $3,200.00
Row 7: INV-1234569   $1,800.00
```
Configure: `Invoice Column`, `Amount Column`, `Header Row Count`

**`SINGLE_CELL`**: All invoice numbers in one cell, delimited
```
Cell B5: "1234567, 1234568, 1234569"
Cell C5: "$5,000.00, $3,200.00, $1,800.00"
```
Configure: `Invoice Delimiter`, `Amount Delimiter`, `Invoice Row`, `Amount Row`

### Regex Pattern Examples
| Pattern | Matches |
|---------|---------|
| `\d{7}` | 7-digit invoice numbers: `1234567` |
| `INV-\d{6}` | Prefixed: `INV-123456` |
| `[A-Z]{2}\d{5}` | Alpha-numeric: `AB12345` |
| `\d{4,8}` | Variable length (4-8 digits) |

Use the **Test** button in the wizard to verify matches before saving.

### Cell Type Colors (Template Builder)
| Color | Type | Description |
|-------|------|-------------|
| Blue | `DATA` | Populated from remittance file |
| Green | `LOOKUP` | Fetched from internal Data Model |
| Yellow | `CALC` | Formula cell |
| White | `BLANK` | Intentionally empty |
| Gray | `LABEL` | Static text / header label |

### Remittance Field Map Keys (DATA cells)
| Key | Value |
|-----|-------|
| `invoice_number` | Extracted invoice number |
| `invoice_amount` | Invoice dollar amount |
| `invoice_date` | Invoice date |
| `payment_date` | Payment/remittance date |
| `payment_amount` | Total remittance amount |
| `check_number` | Check or EFT reference |
| `payor_name` | Payor display name |
| `total_extracted` | Sum of all extracted invoice amounts |
| `variance` | `payment_amount - total_extracted` |
| `invoice_count` | Number of invoices |
| `row_number` | Row index within body (1-based) |

---

## Internal Data Model

### Storage
Data is stored in a hidden sheet `_InternalData_` within the workbook using a flat table structure. Excel's Data Model (Power Pivot) compressed storage is used for 300K+ record performance.

### Refresh Process
1. Open **Internal Data → Refresh Data**
2. Select source file (`.xlsx` or `.csv` export from billing system)
3. Preview schema and first rows
4. Confirm and run refresh
5. Before/after record counts displayed
6. Automatic backup created (rollback available if needed)

### Schema Drift Detection
On each refresh, the engine compares old column names against new ones and reports:
- **Added columns** — new fields available for template mapping
- **Removed columns** — existing template mappings may break

---

## Technical Constraints

- **Offline only** — no cloud dependencies or web APIs
- **Excel 2016+** — Power Query and Data Model are built-in from 2016 onward
- **Windows** — VBA and COM automation require Windows
- **No installation** — runs from a single `.xlsm` file
- **Macro security** — workbook requires macro-enabled (`.xlsm`) format

---

## Error Flags in Output

| Color | Meaning |
|-------|---------|
| Red cell with `#LOOKUP_FAILED` | Invoice not found in Data Model |
| Orange row | Duplicate invoice number detected |
| Red row | Processing error for this invoice |

---

## Audit Log

All processing events are logged to the hidden `AuditLog` sheet:
- Timestamp, action type, payor, source file, detail, status
- View via **Tools → Audit Log** or ribbon button
- Maximum 10,000 entries (oldest auto-purged)

---

## Example: Acme Insurance Co.

**Setup:**
- File type: `.xls` (fake HTML)
- Layout: `SINGLE_CELL`
- Invoice delimiter: `,`
- Regex: `\d{7}`
- Invoice row: 5, Amount row: 5
- Invoice column B, Amount column C

**Remittance input:**
```
Cell B5: "1234567, 1234568, 1234569"
Cell C5: "$5,000.00, $3,200.00, $1,800.00"
```

**Engine output:**
- Detects HTML-as-XLS → converts to proper .xlsx
- Extracts 3 invoices → selects multi-invoice template
- Expands template body × 3, alternating row shading
- Queries Data Model: pulls matter number, client, billing partner
- Applies navy header (#003366 white text), Calibri 10pt body, `$#,##0.00` currency
- Calculates footer totals ($10,000.00)
- Sets print area, freezes header row
- Flags any invoice not in Data Model with red fill
