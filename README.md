# DD1750 Converter

Convert GCSS-Army BOM (Bill of Materials) PDFs to DD Form 1750 Packing Lists.

## Features

- **Automatic Extraction**: Extracts items from GCSS-Army Component Listing / Hand Receipt PDFs
- **Review & Edit Interface**: Verify and modify extracted data before generating
- **Multiple Format Support**: Handles standard GCSS-Army BOMs and EPP format
- **Human-in-the-Loop**: Ensures accuracy with manual review capability
- **Quick Generate**: Option to skip review for experienced users

## Supported BOM Formats

1. **GCSS-Army Component Listing / Hand Receipt**
   - Standard format with LV (Level) column
   - Items marked with LV="B" are extracted as components

2. **Equipment Property Record (EPP)**
   - Alternative format from GCSS-Army

**Note**: Handwritten BOMs are NOT supported. Please obtain clean digital BOMs from GCSS-Army through your supply teams.

## Installation

### Local Development

```bash
# Clone the repository
git clone <your-repo-url>
cd dd1750-converter

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

Visit `http://localhost:8000` in your browser.

### Railway Deployment

1. Create a new Railway project
2. Connect your GitHub repository
3. Railway will automatically detect the Python app and deploy

Or use the Railway CLI:

```bash
railway login
railway init
railway up
```

## Usage

### With Review (Recommended)

1. **Upload**: Select your BOM PDF file
2. **Extract**: Click "Extract Items" to parse the document
3. **Review**: Verify and edit extracted items:
   - Modify descriptions
   - Add/correct NSN numbers
   - Adjust quantities
   - Delete unnecessary items
   - Add missing items
4. **Generate**: Download your DD1750 PDF

### Quick Generate

For experienced users who don't need to review:
1. Use the "Quick Generate" form at the bottom of the page
2. Upload your BOM and optionally a custom DD1750 template
3. Click "Generate" to download directly

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main page with upload form |
| `/upload` | POST | Upload BOM and extract items |
| `/generate` | POST | Generate DD1750 from reviewed items |
| `/quick-generate` | POST | Generate DD1750 without review |
| `/api/formats` | GET | List supported BOM formats |

## Configuration

Environment variables:
- `PORT`: Server port (default: 8000)
- `SECRET_KEY`: Flask secret key for sessions
- `DEBUG`: Enable debug mode ('true'/'false')

## File Structure

```
dd1750_app/
├── app.py              # Flask application
├── dd1750_core.py      # Core extraction and generation logic
├── templates/
│   └── index.html      # Web interface
├── blank_1750.pdf      # DD1750 template
├── requirements.txt    # Python dependencies
├── Procfile           # Process configuration
├── railway.json       # Railway deployment config
└── README.md          # This file
```

## Technical Details

### PDF Processing

- **pdfplumber**: Extracts tables and text from BOM PDFs
- **pypdf**: Merges generated content with DD1750 template
- **reportlab**: Creates PDF overlays with item data

### DD1750 Layout

- Letter size (612 x 792 points)
- 18 rows per page
- Columns: Box No., Contents, Unit of Issue, Initial Operation, Running Spares, Total

## Troubleshooting

### "No items found"
- Ensure you're using a digital GCSS-Army BOM (not handwritten/scanned)
- Try adjusting the start page if your BOM has cover pages
- Check that items have LV="B" designation

### Incorrect quantities
- Review and edit quantities in the review interface
- The extraction looks for "Auth Qty" column values

### Missing NSN
- Add NSN manually in the review interface
- NSN should be 9 digits (NIIN format)

## License

Proprietary - Internal use only.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## Support

For issues or questions, contact your unit's S4 or technology support team.
