#!/usr/bin/env python
# coding: utf-8

from src.utils.constants import REPORT_PATH
from utils.utils import process_all_pl_files, save_summary_with_format


# Параметры
base_directory = r'\\kantar-tns.local\Project\!Methodology\Analytics\!ORG\Финансовая отчетность'
base_directory = r'\\kantar-tns.local\Project\Финансовые_отчеты_Custom\Архив\Предыдущие периоды'
years_to_process = [2024, 2025]
target_pc = 'Analytics'


# Обработка файлов
df_summary = process_all_pl_files(base_directory, years_to_process, target_pc)
df_summary


# Сохранение результатов
save_summary_with_format(df_summary, REPORT_PATH)
print(REPORT_PATH)

