#!/usr/bin/env python
# coding: utf-8

import os

from src.utils.constants import CONFIG_BOT_DIR
from utils.utils import process_all_pl_files


# Параметры
base_directory = r'\\kantar-tns.local\Project\!Methodology\Analytics\!ORG\Финансовая отчетность'
years_to_process = [2024, 2025]
target_pc = 'Analytics'


# Обработка файлов
df_summary = process_all_pl_files(base_directory, years_to_process, target_pc)
df_summary


# Сохранение результатов
out_path = os.path.join(CONFIG_BOT_DIR, 'summary_profit_center.xlsx')
df_summary.to_excel(out_path, index=False)
print(out_path)

