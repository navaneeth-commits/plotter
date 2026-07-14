# Data Canvas

Data Canvas is a data visualization app with an AI-assisted smart dashboard flow:

- frontend: `index.html` + `script.js` + Plotly
- backend: `app.py` with Flask
- AI: Google Gemini via `google-genai`

## Files

- `index.html` - app structure and CDN script imports
- `style.css` - layout and visual styling
- `script.js` - upload, parsing, preview, chart generation, and smart dashboard request logic
- `app.py` - Flask backend for `/eda`

## How to run

1. Install Python packages:
   `pip install flask google-genai`
2. Set your Gemini API key (never commit this — it's read from the environment only):
   - PowerShell: `$env:GEMINI_API_KEY="your_actual_key_here"`
   - bash/zsh: `export GEMINI_API_KEY="your_actual_key_here"`
3. Start the Flask backend:
   `python app.py`
   The server will refuse to start if `GEMINI_API_KEY` isn't set.
4. Open `index.html` in a browser with internet access so the Plotly and SheetJS CDN assets can load.
5. Upload a dataset.
6. Click `Build Smart Dashboard` to generate AI suggestions.

> **Security note:** a previous version of this project had a live API key hardcoded in `app.py` and `test_api.py`. If you ever had that key, revoke/regenerate it in Google AI Studio — hardcoded keys in source are effectively public once committed anywhere.

## Current features

- drag-and-drop or file-picker upload
- automatic parsing for delimited text files    
- spreadsheet parsing using the first worksheet
- inferred numeric, category, date, and id-like column detection
- scatter, line, bar, histogram, box, and pie charts
- optional grouping/color split by a selected column
- AI-powered smart dashboard suggestions through the `/eda` endpoint

## Notes

- Pie charts require a numeric value column.
- For text files, the parser tries common delimiters: comma, tab, semicolon, and pipe.
- Smart Dashboard sends only compact dataset metadata to the backend: column names, inferred types, and the first 5 rows.
- If AI output parsing fails, the backend returns a safe fallback JSON response.
