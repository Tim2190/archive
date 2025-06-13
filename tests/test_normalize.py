import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1]))

from extract_word import normalize_source_id, process_documents


def test_normalize_source_id():
    assert normalize_source_id('Д64-217.doc') == 'D-64-217'
    assert normalize_source_id('АП 90-341.docx') == 'AP-90-341'
    assert normalize_source_id('Кассета АП 60-072 БРИФИНГ РОГОВА.doc') == 'AP-60-072'
    assert normalize_source_id('В 010.doc') == 'V-010'
    assert normalize_source_id('Д124-086 Пресс-конференция.doc') == 'D-124-086'


def test_process_documents_empty(tmp_path):
    out = tmp_path / 'out.csv'
    processed, rows, _ = process_documents(tmp_path, out, limit=10)
    assert processed == 0
    assert rows == 0
    assert not out.exists()
