# Speedy Scandocs

Speedy Scandocs is a desktop tool for law offices that automatically classifies scanned documents (PDFs and JPEGs) by identifying which client they belong to and renaming them with a standardized format: `CLIENT NAME - Document Description.pdf`.

It uses a local AI model (via Ollama/OpenWebUI) to read document text, identify the client, and generate a short description. A fuzzy matching system then maps the AI's client identification to your firm's official client list.

## Features

- **Auto-Process Documents** — Batch-process an entire folder of scanned documents. Each file is analyzed by AI, matched to a client, and renamed automatically.
- **Manual Review** — Documents the AI couldn't confidently identify are flagged as "A-NEEDS REVIEW" for manual assignment.
- **Audit Mode** — After processing, review each result and mark it as correct, wrong client, bad description, etc. Submitting the audit renames flagged files and saves a detailed report with copies of any wrong-client documents.
- **File Mode** — Move renamed files to destination folders (e.g., individual client folders) directly from the app.
- **Document Viewer** — Press **Spacebar** on any selected file to preview it inline. Arrow keys navigate between documents.
- **Excel Reports** — Each audit generates a report with a Summary tab showing completion rates and success metrics.

## Requirements

- **Python 3.10+** (if running from source)
- **Ollama** — Local AI model server (https://ollama.com)
- **OpenWebUI** (optional) — Web interface for Ollama that provides an OpenAI-compatible API
- **Tesseract OCR** — For reading scanned/image-based PDFs (bundled in packaged builds)

### Python Dependencies

Install via `pip install -r requirements.txt`:
- ttkbootstrap
- Pillow
- PyMuPDF (fitz)
- pytesseract
- requests
- openpyxl
- rapidfuzz

## Getting Started

1. **Install and start Ollama** with your chosen model (see recommended settings below).
2. **Prepare your client list** — a plain text file with one client name per line in `LAST, First` format.
3. **Launch the app** — run `python scandocs_tool.py` or open the installed application.
4. **Configure Settings** — set your Scandocs folder path, client list file, API URLs, and processing options.
5. **Click "Auto-Process Documents"** to start batch processing.

## Configuration Settings

All settings are accessible from the **Settings** tab in the app. They are saved to `config.json`.

### Recommended Settings

These are the tested and recommended values for production use:

| Setting | Recommended Value | Description |
|---------|-------------------|-------------|
| **Model** | `ministral-3:8b` | Best balance of speed and accuracy for document classification |
| **Fuzzy Match Threshold** | `0.85` | How closely an AI-identified name must match your client list (0.0-1.0). 0.85 allows minor spelling variations while avoiding false matches |
| **Max OCR Characters** | `8000` | Maximum characters of extracted text sent to the AI. 8000 provides enough context for accurate classification |
| **Max Pages Per Document** | `5` | Number of PDF pages to read. Client names are almost always on the first few pages |
| **Require High Confidence** | `Yes (checked)` | Only auto-rename when the AI is highly confident. Medium/low confidence results go to manual review. Recommended to keep enabled |

### Paths

- **Scandocs Folder** — The folder containing scanned documents to process. The app will read PDFs and JPEGs from this folder.
- **Client List File** — Path to your `client_list.txt` file. One client name per line, in `LAST, First` format.

### API Settings

- **OpenWebUI URL** — URL of your OpenWebUI instance (default: `http://localhost:3000`). The app uses the OpenAI-compatible chat completions endpoint.
- **Ollama URL** — URL of your Ollama instance (default: `http://localhost:11434`). Used as a fallback if OpenWebUI is unavailable.
- **Model** — The AI model to use for classification. Click the refresh button to see models available on your server.
- **API Key** — Only needed if your OpenWebUI instance requires authentication. Leave blank for local setups without auth.

### Processing

- **Fuzzy Match Threshold** — Controls name matching strictness. 1.0 = exact match only, 0.85 = recommended, 0.70 = lenient (more false positives).
- **Max OCR Characters** — Limits the text sent to the AI. Higher values give more context but take longer per file.
- **Max Pages Per Document** — Limits how many pages are extracted from each PDF. Keep at 5 for best performance.
- **Require High Confidence** — When checked, only high-confidence AI results are auto-renamed. Medium-confidence results go to manual review. Recommended: enabled.
- **Skip Already Processed** — When enabled, files that already match the `CLIENT - Description.ext` naming pattern with a recognized client are skipped.

### Reports

- **Report Folder** — Where audit reports are saved. Each audit creates a subfolder named `ScandocsAudit_YYYY-MM-DD_NN` containing the Excel report and copies of any wrong-client documents.
- **Auto-save Report** — Automatically saves a report when batch processing completes.
- **Audit Mode** — Enables the audit panel for reviewing each result before finalizing. Recommended for production use.
- **File Mode** — Enables moving renamed files to destination folders after processing.

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **Spacebar** | Toggle document preview for the selected file |
| **Left / Right Arrow** | Navigate to previous / next document |
| **Double-click** | Open document preview |

## Audit Workflow

1. Run **Auto-Process Documents** on your scanned folder.
2. For each result, review and mark one of:
   - **Correct** — The rename was accurate
   - **Wrong Client Name** — Will be renamed to "A-NEEDS REVIEW" on submit
   - **Bad Description** — Description will be changed to "Scanned Document" on submit
   - **Failed to Identify Client** — The tool couldn't find the client
   - **Should Have Been Flagged** — Should have gone to manual review
3. Click **Submit Audit** to apply corrections, save the report, and copy wrong-client files for review.

## Report Output

Each audit creates a folder in your Reports directory:

```
Reports/
  ScandocsAudit_2025-01-15_01/
    scandocs_report_2025-01-15_143022.xlsx
    SMITH, John - Some Document.pdf        (copy of wrong-client file)
    DOE, Jane - Another Document.pdf       (copy of wrong-client file)
  ScandocsAudit_2025-01-15_02/
    ...
```

The Excel report contains two sheets:
- **Summary** — Overview statistics: total documents, skipped, identified, needs review, audit completion percentage, and success rates.
- **Results** — Detailed per-file results with original name, new name, status, client, description, confidence, and all audit flags.

## Building from Source

### Windows
```bash
cd build/windows
build_windows.bat
```
This creates an installer at `build/windows/Output/SpeedyScandocsSetup.exe`.

### macOS
```bash
pyinstaller build/mac/scandocs_mac.spec --clean
```
This creates `dist/SpeedyScandocs.app`. Use `create-dmg` to package as a DMG for distribution.
