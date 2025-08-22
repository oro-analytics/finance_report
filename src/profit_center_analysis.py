#!/usr/bin/env python
# coding: utf-8

from src.utils.constants import REPORT_PATH
from utils.utils import process_all_pl_files, process_all_x_charge_files, save_summary_with_format


# Параметры
base_directory = r'\\kantar-tns.local\Project\Финансовые_отчеты_Custom\Архив\Предыдущие периоды'
base_directory = r'\\kantar-tns.local\Project\!Methodology\Analytics\!ORG\Финансовая отчетность'
years_to_process = [2025]  #[2024, 2025]
target_pc = 'Analytics'


# Обработка файлов - 'Реализация без НДС', 'Total Direct Costs', 'Total Operating Costs', Operation Profit
print("Обработка файлов - 'Реализация без НДС', 'Total Direct Costs', 'Total Operating Costs', Operation Profit")
print("INFO: Реализация - "
      "Сумма признанной выручки на момент составления P&L отчета, равна сумме контракта при 100% завершенности.")
df_pl_summary = process_all_pl_files(base_directory, years_to_process, target_pc)
row_with_profit_center = process_all_x_charge_files(base_directory, years_to_process, target_pc)
print(row_with_profit_center)

df_summary = df_pl_summary.merge(row_with_profit_center, on=['Год', 'Месяц'], how='left')

# Сохранение результатов
report_path = REPORT_PATH % f"[{','.join([str(x) for x in years_to_process])}]"
save_summary_with_format(df_summary, report_path)
print(report_path)
print(f"Chargeability суммарный за год во вкладке `Chargeability_отчет` файла PL_XX {max(years_to_process)}.xlsx")
