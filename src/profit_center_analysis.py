#!/usr/bin/env python
# coding: utf-8

from src.utils.constants import REPORT_PATH
from utils.utils import process_all_pl_files, process_all_x_charge_files, save_summary_with_format, \
    write_monthly_with_highlights

# Параметры
base_directory = r'\\kantar-tns.local\Project\Финансовые_отчеты_Custom\Архив\Предыдущие периоды'
base_directory = r'\\kantar-tns.local\Project\!Methodology\Analytics\!ORG\Финансовая отчетность'
years_to_process = [2025]  # [2024, 2025]
target_pc = 'Analytics'

# Обработка файлов - 'Реализация без НДС', 'Total Direct Costs', 'Total Operating Costs', Operation Profit
print("Обработка файлов - 'Реализация без НДС', 'Total Direct Costs', 'Total Operating Costs', Operation Profit")
print("INFO: Реализация - "
      "Сумма признанной выручки на момент составления P&L отчета, равна сумме контракта при 100% завершенности.")
df_pl_summary = process_all_pl_files(base_directory, years_to_process, target_pc)
row_with_profit_center_dict, giver_dict, taker_dict = process_all_x_charge_files(base_directory, years_to_process,
                                                                                 target_pc)
# print(row_with_profit_center_dict)

# report_common_name
report_common_name = REPORT_PATH % f"[{','.join([str(x) for x in years_to_process])}]"
# X-charge projects dynamic
report_path = report_common_name.replace("finance_report [", "finance_report_Xcharge_taker [")
result_path = write_monthly_with_highlights(
    dfs_dict=taker_dict,
    output_path=report_path,
    id_col="Номер контракта",
    highlight_first=True,  # первый месяц считаем «появлениями»
    add_flag_column=True
)
print("Готово:", result_path)

#
report_path = report_common_name.replace("finance_report [", "finance_report_Xcharge_giver [")
result_path = write_monthly_with_highlights(
    dfs_dict=giver_dict,
    output_path=report_path,
    id_col="Номер контракта",
    highlight_first=True,  # первый месяц считаем «появлениями»
    add_flag_column=True
)
print("Готово:", result_path)

# Общий
report_path = report_common_name
df_summary = df_pl_summary.merge([row_with_profit_center_dict[k] for k in row_with_profit_center_dict],
                                 on=['Год', 'Месяц'], how='left')

save_summary_with_format(df_summary, report_path)
print(report_path)
print(f"Chargeability суммарный за год во вкладке `Chargeability_отчет` файла PL_XX {max(years_to_process)}.xlsx")
