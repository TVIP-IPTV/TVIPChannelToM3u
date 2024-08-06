import sys
import zipfile
import xml.etree.ElementTree as ET
import os
import pandas as pd
import shutil
from loguru import logger
from config import Config

cfg = Config()

logger.remove()
logger.add(sys.stderr, format='<white>{time:HH:mm:ss}</white>'
                           ' | <level>{level: <8}</level>'
                           ' | <cyan>{line}</cyan>'
                           ' - <white>{message}</white>')

def replace_in_xml(file_path, replacements):
    tree = ET.parse(file_path)
    root = tree.getroot()
    for elem in root.iter():
        if elem.text in replacements:
            elem.text = replacements[elem.text]
    tree.write(file_path)

def process_excel_zip(zip_path, extract_path='extracted_files'):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
    except FileNotFoundError:
        logger.error(f'Файл {zip_path} не найден')
        raise
    
    replacements = {
        '[object Object]': '',
        '[object Object],[object Object]': '',
        '[object Object],[object Object],[object Object]': '',
        '[object Object],[object Object],[object Object],[object Object]': '',
        '[object Object],[object Object],[object Object],[object Object],[object Object]': ''
    }
    
    for foldername, _, filenames in os.walk(os.path.join(extract_path, 'xl/worksheets')):
        for filename in filenames:
            if filename.endswith('.xml'):
                file_path = os.path.join(foldername, filename)
                replace_in_xml(file_path, replacements)
    
    fixed_excel_path = 'fixed_tv_test.xlsx'
    with zipfile.ZipFile(fixed_excel_path, 'w') as zip_ref:
        for foldername, _, filenames in os.walk(extract_path):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zip_ref.write(file_path, os.path.relpath(file_path, extract_path))
    
    def cleanup_directory(directory):
        if os.path.exists(directory):
            shutil.rmtree(directory)

    cleanup_directory('extracted_files')

    return fixed_excel_path

def convert_to_m3u(excel_path, m3u_path):
    df = pd.read_excel(excel_path)
    required_columns = ['enabled', 'displayName', 'url', 'logoUrl', 'textName']
    
    for column in required_columns:
        if column not in df.columns:
            logger.error(f'Столбец {column} не найден в Excel файле')
            raise ValueError(f'Столбец {column} не найден в Excel файле')

    with open(m3u_path, 'w', encoding='utf-8') as m3u_file:
        for index, row in df.iterrows():
            try:
                if row['enabled'] and isinstance(row['displayName'], str) and isinstance(row['url'], str) and isinstance(row['textName'], str):
                    logo = row['logoUrl'] if isinstance(row['logoUrl'], str) else ''
                    m3u_file.write(f'#EXTINF:-1 tvg-id="{row["textName"]}" tvg-logo="{logo}",{row["displayName"]}\n')
                    m3u_file.write(f'{row["url"]}\n')
            except (ValueError, TypeError) as e:
                logger.warning(f'В работе процесса {index} произошла ошибка: {e}')
                continue

    logger.success('Конвертация завершена успешно')

if __name__ == '__main__':
    try:
        input_name = input('Введите название xlsx файла с каналами: ')
        output_name = input('Введите название m3u файла (итоговый плейлист): ')

        if cfg.is_alpha_version():
            logger.warning(f'Вы используете Alpha версию ({cfg.__version__}). В работе программы могут быть ошибки, если вы нашли ошибку, оставьте issues в репозитории')

        fixed_excel_path = process_excel_zip(input_name)
        convert_to_m3u(fixed_excel_path, output_name)
        os.remove(fixed_excel_path)
    except Exception as e:
        logger.error(f'Произошла ошибка: {e}')
        sys.exit(1)
