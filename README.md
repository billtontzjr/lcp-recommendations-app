# LCP Recommendations App

A web application that generates Life Care Plan (LCP) Recommendations documents from Excel workbooks. Built for Precision Life Care Planning by Dr. William Tontz, MD, CLCP.

## Overview

This app takes a Master Workbook (.xlsm/.xlsx) containing patient information and selected care items, calculates lifetime costs, and generates a professionally formatted Word document (.docx) with:

- Title page with patient details
- Appendix A summary table with lifetime projected costs
- Individual section pages for each care category (Physicians, Therapies, Medications, etc.)
- Rationale and Sources tables for each section

## Tech Stack

- **Backend:** Python 3.11, Flask
- **Document Generation:** python-docx, openpyxl
- **Database:** Supabase (PostgreSQL) - optional
- **Deployment:** Render.com
- **Frontend:** Vanilla HTML/CSS/JavaScript

## Project Structure

```
lcp-recommendations-app/
├── app/
│   ├── __init__.py
│   ├── main.py                 # Flask app factory & entry point
│   ├── config.py               # Environment configuration
│   │
│   ├── routes/
│   │   ├── __init__.py
│   │   ├── api.py              # API endpoints (/api/generate, /api/preview)
│   │   └── health.py           # Health check for Render
│   │
│   ├── services/
│   │   ├── __init__.py
│   │   ├── workbook_parser.py  # Excel file parsing
│   │   ├── cost_calculator.py  # Cost calculations & aggregation
│   │   ├── document_generator.py # Word document generation
│   │   └── supabase_client.py  # Database integration (optional)
│   │
│   └── utils/
│       ├── __init__.py
│       ├── frequency_parser.py # Parses frequency strings ("2x/year", "every 5 years")
│       └── currency.py         # Currency formatting utilities
│
├── templates/
│   └── index.html              # Upload page UI
│
├── static/
│   ├── css/
│   │   └── style.css
│   └── js/
│       └── app.js              # Frontend file upload & preview logic
│
├── tests/
│   └── __init__.py
│
├── requirements.txt
├── Procfile                    # Render/Heroku start command
├── render.yaml                 # Render deployment config
├── .env.example                # Environment variables template
└── .gitignore
```

## How It Works

### 1. Workbook Parsing (`workbook_parser.py`)

Reads the uploaded Excel workbook and extracts:

- **Patient Info sheet** (Row 4-14, Column E):
  - Date of Report, Patient Name, DOB, Age, Date of Injury
  - Life Expectancy, Age Initiated, Geographic Multiplier
  - City/State, Zipcode, Referring Attorney

- **Master sheet** (Starting Row 6):
  - Column A: Check (Boolean) - only checked items are included
  - Column B: Main Category
  - Column C: Item name
  - Column D: Subcategory
  - Column E: Service Description
  - Column F: Code Type (PFR, APC, DRG)
  - Column G: CPT Code(s)
  - Column H: Cost (pre-calculated)
  - Column I: Frequency
  - Column J: Source
  - Column K: Rationale

- **PFR sheet**: Professional fee lookup (CPT Code → P75 Price)
- **APC sheet**: Facility fee lookup (CPT Code → Facility Charge)

### 2. Cost Calculation (`cost_calculator.py`)

For each selected item:

1. Parse frequency string to determine if annual or one-time
2. Look up costs from PFR/APC tables if not provided
3. Apply geographic multiplier to facility fees
4. Calculate:
   - `annual_cost = unit_cost × frequency_multiplier` (for recurring items)
   - `one_time_cost = unit_cost` (for one-time items)
   - `lifetime_annual = annual_cost × life_expectancy`
   - `grand_total = lifetime_annual + one_time_cost`

**Frequency Patterns Supported:**
- `"2x/year"`, `"3 times per year"`, `"monthly"` → Annual multiplier
- `"every 5 years"`, `"every 8-10 years"` → Fractional annual
- `"one time"`, `"one-time"` → One-time cost
- `"24 visits every 5 years"` → Visits ÷ Years

### 3. Document Generation (`document_generator.py`)

Generates a Word document with:

- **Title Page**: Patient name, DOB, DOI, Life Expectancy, Prepared by credentials
- **Appendix A**: Summary table with all categories and totals
- **Section Pages**: One per category with:
  - 7-column data table (Item, Age Init, Until Age, Cost, Frequency, Annual, One-Time)
  - Section totals table
  - Rationale table
  - Sources table

**Table Formatting:**
- Style: Table Grid with bold borders (sz=12)
- Gray headers: `#D3D3D3`
- Font: Times New Roman, 10pt
- Cell padding: 70 twips top/bottom
- Empty gray separator rows between sections

## API Endpoints

### `POST /api/generate`

Upload workbook and generate LCP document.

**Request:**
```
Content-Type: multipart/form-data
- file: Master Workbook (.xlsx/.xlsm) [required]
- medical_summary: Medical Summary (.docx) [optional]
```

**Response:** Downloads generated .docx file

### `POST /api/preview`

Preview cost calculations without generating document.

**Request:**
```
Content-Type: multipart/form-data
- file: Master Workbook (.xlsx/.xlsm)
```

**Response:**
```json
{
  "patient_info": {
    "patient_name": "John Doe",
    "date_of_birth": "1980-01-15",
    "life_expectancy": 35.5
  },
  "totals": {
    "total_annual": 15000.00,
    "total_one_time": 25000.00,
    "lifetime_annual": 532500.00,
    "grand_total": 557500.00
  },
  "categories": {
    "Physicians": { "annual_cost": 5000, "one_time_cost": 0, "item_count": 3 }
  },
  "item_count": 25
}
```

### `GET /health`

Health check endpoint for Render deployment.

## Local Development

### Prerequisites
- Python 3.11+
- pip

### Setup

```bash
# Clone repository
git clone https://github.com/billtontzjr/lcp-recommendations-app.git
cd lcp-recommendations-app

# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Copy environment template
cp .env.example .env
# Edit .env with your settings (Supabase is optional)

# Run development server
python app.py
```

Open http://localhost:5000

### Running with Gunicorn (Production-like)

```bash
gunicorn app.main:app --bind 0.0.0.0:5000
```

## Deployment (Render)

### Option 1: Auto-deploy from GitHub

1. Push code to GitHub
2. Connect repo to Render
3. Render auto-detects `Procfile` and deploys

### Option 2: Manual Configuration

- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `gunicorn app.main:app --bind 0.0.0.0:$PORT`
- **Health Check Path:** `/health`

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `SECRET_KEY` | Yes | Flask secret key (auto-generated on Render) |
| `SUPABASE_URL` | No | Supabase project URL |
| `SUPABASE_KEY` | No | Supabase anon key |
| `PORT` | Auto | Set by Render |

## Supabase Integration (Optional)

If configured, the app saves:
- Case records with patient info and totals
- Individual care items linked to cases
- Generated documents to Supabase Storage

### Database Schema

See `supabase_setup.sql` for table definitions:
- `cases` - Patient info and cost totals
- `case_items` - Individual care items
- `documents` - Document metadata

## Customization

### Adding New Frequency Patterns

Edit `app/utils/frequency_parser.py`:

```python
FREQUENCY_MULTIPLIERS = {
    "yearly": 1.0,
    "monthly": 12.0,
    # Add new patterns here
}
```

### Modifying Document Layout

Edit `app/services/document_generator.py`:

- `add_title_page()` - Title page content
- `add_appendix_a()` - Summary table structure
- `add_section_page()` - Section table columns and formatting
- `set_bold_borders()` - Table border styling
- `GRAY_BACKGROUND` - Header background color

### Changing Table Column Widths

In `document_generator.py`, modify the `col_widths` arrays:

```python
# Appendix A table
col_widths = [Inches(1.5), Inches(1.2), Inches(1.5), Inches(2.5), Inches(1.5)]

# Section tables
col_widths = [Inches(2), Inches(1), Inches(1), Inches(1), Inches(1.5), Inches(1.5), Inches(1.5)]
```

## Troubleshooting

### "No items selected in workbook"
- Ensure Column A (Check) has `TRUE` for items to include
- Check that data starts at Row 6 in Master sheet

### "Patient name is required"
- Verify Patient Info sheet exists
- Check Row 5, Column E contains patient name

### Document formatting issues
- Ensure python-docx version >= 0.8.11
- Check that all table cells have content (empty strings, not None)

## License

Proprietary - Precision Life Care Planning

## Contact

Dr. William Tontz, Jr., MD, CLCP
Precision Life Care Planning
