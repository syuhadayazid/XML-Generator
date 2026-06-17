# Python XML Generator

Generate sample XML from Excel mapping files with auto-detected XML path columns and SAMPLE_VALUE mode.

## Features

- **Auto-detect XML path column** – Scans your spreadsheet and finds the column containing XML/XPath-like paths
- **Browser-based UI** – Use the web interface without installing Python
- **Python CLI** – Run from terminal for automation/batch processing
- **SAMPLE_VALUE mode** – Fill all generated XML values with `SAMPLE_VALUE` for testing
- **Group same fields** – Sibling elements with the same tag name are grouped together
- **No backend required** – Browser UI runs fully in-browser; Python version is self-contained

## Quick Start

### Browser UI (Recommended)

1. Open [docs/python_xml_generator_ui.html](docs/python_xml_generator_ui.html) in your browser
2. Upload your Excel mapping file (.xlsx)
3. Auto-detect finds the XML path column automatically
4. Click **Generate XML**
5. Download or copy the generated XML

### Python CLI

#### Installation

1. Install Python 3.8+
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

#### Basic Usage

```bash
python python_xml_generator/generate_sample_xml.py input.xlsx --sample-only -o output.xml
```

#### Options

```
--worksheet SHEET_NAME       Specify which sheet to use (default: first sheet)
--path-column-ref B          Use a fixed column (e.g., B, C, or 1, 2, 3)
--sample-only                Force all values to SAMPLE_VALUE
-o, --output FILE            Output file path (default: sample.xml)
```

#### Examples

Auto-detect XML column, force SAMPLE_VALUE:
```bash
python python_xml_generator/generate_sample_xml.py mapping.xlsx --sample-only -o sample.xml
```

Use column B explicitly, specific worksheet:
```bash
python python_xml_generator/generate_sample_xml.py mapping.xlsx \
  --path-column-ref B \
  --worksheet "X12 214-4010_to_IDM" \
  --sample-only \
  -o output.xml
```

## Excel Format Expected

Your Excel file should have:
- **Column B (or detected)**: XML/XPath-like paths  
  Example: `/px:NotifyShipment/px:DataArea/px:Shipment/px:ShipmentHeader/px:ID[@typeCode='ABC']`
- **Optional other columns**: Mapping guidance, comments, etc.

## How It Works

1. **Reads Excel** – Loads your workbook and detects the XML path column
2. **Parses paths** – Normalizes XPath-like paths and splits into segments
3. **Builds XML tree** – Creates nested XML elements with attributes from predicates
4. **Groups siblings** – Reorders same-named elements together
5. **Outputs XML** – Saves formatted, indented XML or displays in browser

## Column Auto-Detection

The tool scans for columns containing path-like text:
- Must contain `/` (forward slashes)
- Must have at least 2 segments separated by `/`
- Must match XML element naming patterns
- Requires ≥3 matches in sample rows to be confident

If auto-detect fails, use `--path-column-ref` to specify the column manually (e.g., `B` or `2`).

## Project Structure

```
.
├── docs/
│   ├── index.html                      # Main docs page
│   ├── python_xml_generator_ui.html    # Browser UI for XML generation
│   └── X12-*.sef                       # Sample SEF schema files
├── python_xml_generator/
│   └── generate_sample_xml.py          # Main Python script
├── requirements.txt                    # Python dependencies
└── README.md                           # This file
```

## Requirements

- Python 3.8+
- openpyxl (for Excel reading)

## License

MIT

## Support

For issues or feature requests, please open an issue in this repository.
