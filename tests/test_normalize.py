import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1]))

from extract_word import normalize_source_id, process_documents


def test_normalize_source_id():
    assert normalize_source_id('Д64-217.doc') == 'D-64-217'
    assert normalize_source_id('АП 90-341.docx') == 'AP-90-341'


def test_process_documents_empty(tmp_path):
    out = tmp_path / 'out.csv'
    processed, rows, _ = process_documents(tmp_path, out, limit=10)
    assert processed == 0
    assert rows == 0
    assert not out.exists()
