import streamlit as st
from pathlib import Path
from extract_word import process_documents
import pandas as pd

st.title('Парсер монтажных листов')

input_dir = st.text_input(
    'Путь к папке с Word-документами', str(Path('input_docs').resolve())
)
output_name = st.text_input(
    'Имя выходного CSV', str(Path('output/result.csv').resolve())
)
limit = st.number_input('Лимит на количество файлов', min_value=1, max_value=500, value=500, step=1)

if st.button('Запустить обработку'):
    path = Path(input_dir).expanduser()
    if not path.exists():
        st.error('Указанная папка не существует')
    else:
        output_path = Path(output_name).expanduser()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        st.write(f'Сохраняем результат в: {output_path.resolve()}')
        processed, rows, log = process_documents(path, output_path, limit)
        st.success(f'Файлов обработано: {processed}, строк получено: {rows}')
        with open(output_path, 'rb') as f:
            st.download_button('Скачать CSV', f.read(), file_name=output_path.name, mime='text/csv')
        if st.checkbox('Показать результат'):
            df = pd.read_csv(output_path, delimiter=';')
            st.dataframe(df)
        st.text_area('Лог', '\n'.join(log), height=200)
