import argparse
import csv
import re
from pathlib import Path

try:
    import win32com.client  # type: ignore
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

TIME_PATTERN = re.compile(r'\b\d{1,2}[:.]\d{2}[:.]\d{2}\b')


def normalize_source_id(filename: str) -> str:
    name = Path(filename).stem.upper()
    for word in EXCLUDE_WORDS:
        name = re.sub(word.upper(), '', name)
    name = name.replace('_', '').replace(' ', '')
    for cyr, lat in TRANSLIT_MAP.items():
        name = name.replace(cyr, lat)
    match = re.match(r'([A-Z]+)(.*)', name)
    if match:
        letters, rest = match.groups()
        rest = rest.lstrip('-')
        if rest and not rest.startswith('-'):
            rest = '-' + rest
        return f"{letters}{rest}"
    return name


def convert_doc_to_docx(path: Path) -> Path:
    if win32com is None:
        raise RuntimeError('win32com is required to handle .doc files')
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(str(path))
    docx_path = str(path) + 'x'
    doc.SaveAs(docx_path, FileFormat=16)
    doc.Close()
    word.Quit()
    return Path(docx_path)


def parse_docx_tables(path: Path):
    from docx import Document  # imported here to avoid dependency when unused
    doc = Document(str(path))
    records = []
    for table in doc.tables:
        if not table.rows:
            continue
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            times = []
            desc_parts = []
            for text in cells:
                found = TIME_PATTERN.findall(text)
                if found:
                    times.extend(found)
                cleaned = TIME_PATTERN.sub('', text).strip()
                if cleaned and not cleaned.isdigit():
                    desc_parts.append(cleaned)
            if not times:
                continue
            if len(times) >= 2:
                tc = f"{times[0]} – {times[1]}"
            else:
                tc = times[0]
            desc = ' '.join(desc_parts)
            for junk in JUNK_PHRASES:
                desc = desc.replace(junk, '').strip()
            if not desc:
                continue
            records.append((tc, desc))
    return records


def write_csv(path: Path, rows):
    exists = path.exists()
    with open(path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';')
        if not exists:
            writer.writerow(['source_id', 'timecode', 'description'])
        for r in rows:
            writer.writerow(r)


def process_documents(input_dir: Path, output_csv: Path, limit: int = 500):
    """Process Word documents and save episodes to CSV.

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

    for file in files:
        try:
            source_id = normalize_source_id(file.name)
            target = file
            if file.suffix.lower() == '.doc':
                target = convert_doc_to_docx(file)
            rows = parse_docx_tables(target)
            if not rows:
                continue
            processed += 1
            row_count += len(rows)
            rows = [(source_id, tc, desc) for tc, desc in rows]
            write_csv(output_csv, rows)
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
