# archive

Utility to extract episode tables from Word documents into a single CSV file.

## Usage

1. Put all `.doc` and `.docx` files inside a directory.
2. Run the extractor specifying that directory:
   
   ```bash
   python extract_word.py /path/to/folder -o result.csv
   ```

Each run processes up to 500 files (configurable with `--limit`) and appends
rows to the CSV. A log file `process.log` lists processed files and any errors.

### Web interface

Run a small GUI with Streamlit:

```bash
streamlit run app.py
```

The interface lets you choose the folder with Word documents, specify the
output CSV path and the number of files to process, then shows progress and a
download link.

Dependencies:

- `python-docx`
- `pywin32` (only on Windows, for `.doc` files)
