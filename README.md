# archive

Utility to extract episode tables from Word documents into a single CSV file.

## Usage

1. Create a folder for your source files, for example `input_docs`, and place
   all `.doc` and `.docx` files there.
2. Run the extractor specifying that directory:
   
   ```bash
   python extract_word.py /path/to/folder -o result.csv
   ```

Each run processes up to 500 files (configurable with `--limit`) and appends
rows to the CSV. The resulting file has three columns: `source_id`, `date`
and `description`. A log file `process.log` lists processed files and any
errors.

### Web interface

Run a small GUI with Streamlit:

```bash
streamlit run app.py
```

The interface lets you choose the folder with Word documents, specify the
output CSV path and the number of files to process, then shows progress and a
download link. By default it looks for `input_docs` and writes to
`output/result.csv` relative to the project directory.

### Requirements

- Python 3.9+
- `python-docx`
- `pywin32` (only on Windows, for `.doc` files)
- Microsoft Word installed if `.doc` support is needed.

The script uses COM automation to convert `.doc` files to `.docx`. If you see
an error like `-2147221008, 'Не был произведен вызов CoInitialize.'`, ensure
that `pywin32` is installed and the program is run on Windows with Microsoft
Word available. The utility initializes COM automatically when converting.
