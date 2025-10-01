# Excel to PNG Screenshot Tool

A Python tool that converts Excel spreadsheet sheets into high-quality PNG images using Playwright and Chromium.

## Overview

This tool reads Excel files (`.xlsx`), converts each sheet to HTML, and uses Playwright's headless Chromium browser to render and capture clean screenshots of the tables. Perfect for generating visual reports, documentation, or sharing spreadsheet data as images.

## Features

- ✅ Export all sheets from an Excel file automatically
- ✅ High-resolution output (configurable DPI/device pixel ratio)
- ✅ Clean table rendering with proper borders and styling
- ✅ Handles complex Excel formatting
- ✅ Progress tracking and error handling
- ✅ Automatic sheet name sanitization for file names

## Requirements

- **Python 3.8+**
- **Playwright** (automatically installs Chromium browser on first setup)

## Installation

1. **Clone or download this repository**

2. **Install Python dependencies:**

   ```bash
   pip install -r requirements.txt
   pip install playwright openpyxl xlsx2html
   ```

3. **Install Playwright browsers (Chromium):**

   ```bash
   playwright install chromium
   ```

   > **Note:** Playwright will download and install Chromium (~170MB) automatically. This is a self-contained browser that doesn't interfere with your system's Chrome/Chromium installation.

## Usage

### Basic Usage

Edit the configuration in `src/app.py`:

```python
EXCEL_FILE = "./src/your_file.xlsx"
OUTPUT_DIR = "./output"
DPI = 2  # 1=normal, 2=high-res, 3=ultra high-res
```

Then run:

```bash
python src/app.py
```

### As a Module

You can also import and use the function programmatically:

```python
from src.app import export_excel_sheets_to_images

export_excel_sheets_to_images(
    xlsx_path="path/to/your/file.xlsx",
    output_dir="./output",
    dpi=2
)
```

### Configuration Options

| Parameter    | Type     | Default      | Description                                                        |
| ------------ | -------- | ------------ | ------------------------------------------------------------------ |
| `xlsx_path`  | str/Path | required     | Path to your Excel file                                            |
| `output_dir` | str/Path | `"./output"` | Directory where PNG files will be saved                            |
| `dpi`        | int      | `2`          | Device pixel ratio (1=normal, 2=retina/high-res, 3=ultra high-res) |

## Output

The tool creates PNG files in the output directory, one for each sheet:

```
output/
├── Sheet1.png
├── Sales_Report.png
├── Data_Analysis.png
└── Summary.png
```

Sheet names with spaces or special characters are automatically sanitized (e.g., `Sales/Report` becomes `Sales_Report.png`).

## How It Works

1. **Load Excel**: Uses `openpyxl` to read the Excel file and enumerate all sheets
2. **Convert to HTML**: Uses `xlsx2html` to convert each sheet to HTML format
3. **Render in Browser**: Playwright launches a headless Chromium browser
4. **Apply Styling**: Injects CSS for clean, professional table rendering
5. **Capture Screenshot**: Takes a screenshot of just the table element (not the entire page)
6. **Save as PNG**: Outputs high-resolution PNG files

## Troubleshooting

### "Playwright is not installed"

Run: `playwright install chromium`

### "Chromium executable not found"

Ensure you ran `playwright install chromium` after installing the playwright package.

### Low-quality images

Increase the `DPI` parameter to `3` for ultra high-resolution output.

### Missing dependencies

Install all requirements: `pip install playwright openpyxl xlsx2html`

## Technical Details

- **Core Technology**: Playwright (headless browser automation)
- **Browser**: Chromium (installed automatically by Playwright)
- **Excel Parsing**: openpyxl
- **HTML Conversion**: xlsx2html
- **Image Format**: PNG with configurable resolution

## License

This project is provided as-is for personal or commercial use.

## Contributing

Feel free to submit issues or pull requests for improvements.
