import argparse
import csv
import re
from pathlib import Path

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    win32com = None


TRANSLIT_MAP = {
    'АП': 'AP',
    'Д': 'D',
    'Ф': 'F',
    'Н': 'H',
}

EXCLUDE_WORDS = [
    'кассета',
    'обращение',
    'копия',
    'монтажный лист',
]

JUNK_PHRASES = ['обрыв', 'таймкод сбит']
TECH_PREFIXES = ['НОМЕР', 'ТАЙМК', 'НАЧАЛ', 'КОНЕЦ', 'TIME']

TIME_PATTERN = re.compile(r'\b\d{1,2}[:.]\d{2}[:.]\d{2}\b')
DATE_PATTERN = re.compile(r'\b\d{1,2}[./]\d{1,2}[./]\d{4}\b')


def format_description(text: str) -> str:
    """Return text with each sentence on a new line."""
    return re.sub(r'(?<=[.!?])\s+', '\n', text).strip()


def normalize_source_id(filename: str) -> str:
    """Return normalized cassette ID based on filename without regex."""
    name = Path(filename).stem.upper()
    for word in EXCLUDE_WORDS:
        name = name.replace(word.upper(), ' ')

    name = name.replace('_', ' ')
    cleaned = ''.join(ch if ch.isalnum() or ch in '- ' else ' ' for ch in name)
    parts = [p for p in cleaned.split() if p]

    letters = ''
    digits = ''
    for part in parts:
        part_letters = ''.join(c for c in part if c.isalpha())
        part_digits = ''.join(c for c in part if c.isdigit() or c == '-')
        if not letters and part_letters:
            letters = part_letters
            if part_digits:
                digits = part_digits
                break
        elif letters and not digits and part_digits:
            digits = part_digits
            break

    if not letters and parts:
        first = parts[0]
        letters = ''.join(c for c in first if c.isalpha())
        digits = ''.join(c for c in first if c.isdigit() or c == '-')

    for cyr, lat in TRANSLIT_MAP.items():
        letters = letters.replace(cyr, lat)

    if letters.startswith('В'):
        letters = 'B' + letters[1:]

    letters = letters.strip('-')
    digits = digits.strip('-')
    return f"{letters}-{digits}".strip('-')


def convert_doc_to_docx(path: Path) -> Path:
    if win32com is None:
        raise RuntimeError('win32com is required to handle .doc files')
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(str(path))
        docx_path = str(path) + 'x'
        doc.SaveAs(docx_path, FileFormat=16)
        doc.Close()
        word.Quit()
        return Path(docx_path)
    finally:
        pythoncom.CoUninitialize()


def parse_docx_tables(path: Path):
    """Return date string and one meaningful description from tables."""
    from docx import Document  # imported here to avoid dependency when unused
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(path))
    date = None
    records: list[str] = []

    for element in doc.element.body.iterchildren():
        if isinstance(element, CT_P):
            if date is None:
                paragraph = Paragraph(element, doc)
                match = DATE_PATTERN.search(paragraph.text)
                if match:
                    date = match.group().replace('/', '.')
        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            if not table.rows:
                continue
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                desc_parts = []
                for text in cells:
                    cleaned = TIME_PATTERN.sub('', text).strip()
                    for junk in JUNK_PHRASES:
                        cleaned = cleaned.replace(junk, '').strip()
                    if cleaned and not cleaned.isdigit():
                        desc_parts.append(cleaned)
                desc = ' '.join(desc_parts).strip()
                if desc:
                    clean_line = ' '.join(desc.split())
                    if (len(clean_line) > 100 and not any(
                            clean_line.upper().startswith(p) for p in TECH_PREFIXES)):
                        records.append(format_description(clean_line))
                        return date, records
    return date, records


def write_csv(path: Path, rows):
    exists = path.exists()
    with open(path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';')
        if not exists:
            writer.writerow(['source_id', 'date', 'description'])
        for r in rows:
            writer.writerow(r)


def process_documents(input_dir: Path, output_csv: Path, limit: int = 500):
    """Process Word documents and save episode descriptions to CSV.

    Parameters
    ----------
    input_dir: Path
        Directory with Word files.
    output_csv: Path
        Destination CSV path.
    limit: int
        Maximum number of files to process.

    Returns
    -------
    tuple[int, int, list[str]]
        Files processed, rows written and log lines.
    """
    files = sorted(input_dir.glob('*.doc*'))[:limit]
    processed = 0
    row_count = 0
    log_lines: list[str] = []
    seen_ids: set[str] = set()

    for file in files:
        try:
            source_id = normalize_source_id(file.name)
            if source_id in seen_ids:
                log_lines.append(f'Skipped duplicate {file.name}')
                continue
            target = file
            if file.suffix.lower() == '.doc':
                target = convert_doc_to_docx(file)
            date, descriptions = parse_docx_tables(target)
            if not descriptions:
                continue
            processed += 1
            row_count += len(descriptions)
            rows = [(source_id, date or '', desc) for desc in descriptions]
            write_csv(output_csv, rows)
            seen_ids.add(source_id)
            log_lines.append(f'Processed {file.name}')
        except Exception as exc:  # pragma: no cover - execution errors
            log_lines.append(f'{file.name}: {exc}')

    return processed, row_count, log_lines


def main():
    parser = argparse.ArgumentParser(description='Extract episode tables from Word docs into CSV')
    parser.add_argument('input_dir', help='Directory with Word files')
    parser.add_argument('-o', '--output', default='result.csv', help='Output CSV path')
    parser.add_argument('--limit', type=int, default=500, help='Files to process in one run')
    parser.add_argument('--log-file', default='process.log', help='Path to log file')
    args = parser.parse_args()

    processed, rows, log_lines = process_documents(Path(args.input_dir), Path(args.output), args.limit)
    with open(args.log_file, 'w', encoding='utf-8') as lf:
        lf.write('\n'.join(log_lines))
    print(f'Files processed: {processed}, rows written: {rows}')


if __name__ == '__main__':
    main()
