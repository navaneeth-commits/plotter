Developed by Navaneeth Kumar Reddy
# Data Canvas

Data Canvas is a lightweight browser-based visualization site that lets you:

- upload `csv`, `txt`, `xls`, or `xlsx` files
- inspect rows and inferred column types
- choose chart types and axis mappings
- generate plots directly in the browser

## Files

- `index.html` - app structure and CDN script imports
- `style.css` - layout and visual styling
- `script.js` - upload, parsing, preview, and chart generation logic

## How to run

Open `index.html` in a browser with internet access so the Plotly and SheetJS CDN assets can load.

## Current features

- drag-and-drop or file-picker upload
- automatic parsing for delimited text files
- spreadsheet parsing using the first worksheet
- inferred numeric vs text column detection
- scatter, line, bar, histogram, box, and pie charts
- optional grouping/color split by a selected column

## Notes

- Pie charts require a numeric value column.
- For text files, the parser tries common delimiters: comma, tab, semicolon, and pipe.
- Charts are rendered client-side; no backend is required for the MVP.
