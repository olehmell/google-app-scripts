# google-app-scripts
Usefull scripts for using in Google Docs, Sheets

## Google Docs Highlights Extractor
Extract highlighted text from Google Docs into color-organized columns in Google Sheets. Designed for Value Proposition Canvas but useful for any color-coded text extraction.

### Quick Setup

- Open Google Sheets → Extensions → Apps Script
- Create new script and paste the code

### Usage
- Put Google Docs links in Column A of your spreadsheet
- Run processMultipleDocsHighlights function from Apps Script
- Get results in separate sheets (one per document)

### Features

- Processes multiple documents
- Auto-detects all highlight colors
- Creates color-coded columns
- Updates existing sheets with same names

### Tips

- Ensure you have edit access to all docs (and it is native Google Docs file, not DOCX file opened with Google Docs. If it is DOCX, just click Save as Google Docs)
- Each link goes in separate row in Column A
