# Universal Dataset → Excel Website

A web interface for the Universal Dataset Downloader v5 engine.

## User experience

The user only needs to:

1. Open the website.
2. Paste a public HTTP/HTTPS dataset URL.
3. Click **Convert to Excel**.
4. Click **Download Excel**.

No Python, terminal, package installation, or command is required on the user's computer.

## Features retained from the converter engine

- GitHub blob → raw URL conversion
- GitHub raw URLs
- Hugging Face dataset URLs
- Direct file URLs
- CSV / TSV / TXT
- JSON / JSONL / JSONL.GZ
- Parquet
- Excel
- HTML tables
- Markdown tables
- ZIP / TAR / TAR.GZ / TGZ
- SQLite
- XML
- YAML
- DOCX
- PPTX
- PDF
- NumPy / MAT / HDF / ARFF where the required decoder is available
- Unknown extensions are detected from file content
- Unknown binary files get a metadata/preview Excel fallback
- Excel files are split into sheets when row count requires it
- Duplicate columns and control characters are cleaned
- Retry handling for common temporary HTTP errors
- Public-access limitation is respected; login/CAPTCHA/private files cannot be bypassed

## Run on a server

The server needs Python and the packages in `requirements.txt`.

```bash
pip install -r requirements.txt
python app.py
```

Then open:

`http://localhost:5000`

## Deploy

`render.yaml` and `Procfile` are included for simple deployment on a Python web host such as Render.

For production, use a persistent/temporary job cleanup policy and configure resource limits appropriate for the size of datasets you expect to process.
