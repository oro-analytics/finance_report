#!/usr/bin/env python
# coding: utf-8
from typing import Optional

import pandas as pd

from src.utils.constants import REPORT_PATH
from utils.utils import process_all_pl_files, process_all_x_charge_files, save_summary_with_format, \
    write_monthly_with_highlights

# Параметры
base_directory = r'\\kantar-tns.local\Project\Финансовые_отчеты_Custom\Архив\Предыдущие периоды'
base_directory = r'\\kantar-tns.local\Project\!Methodology\Analytics\!ORG\Финансовая отчетность'
# base_directory = r'\\kantar-tns.local\Project\!Methodology\Analytics\!ORG\Финансовая отчетность\2025\Test'
years_to_process = [2025]  # [2024, 2025]
target_pc = 'Analytics'

# Обработка файлов - 'Реализация без НДС', 'Total Direct Costs', 'Total Operating Costs', Operation Profit
print("Обработка файлов - 'Реализация без НДС', 'Total Direct Costs', 'Total Operating Costs', Operation Profit")
print("INFO: Реализация - "
      "Сумма признанной выручки на момент составления P&L отчета, равна сумме контракта при 100% завершенности.")
df_pl_summary, pl_projects_dict = process_all_pl_files(base_directory, years_to_process, target_pc)
row_with_profit_center_dict, giver_dict, taker_dict = process_all_x_charge_files(base_directory, years_to_process,
                                                                                 target_pc)
# print(row_with_profit_center_dict)

# report_common_name
report_common_name = REPORT_PATH % f"[{','.join([str(x) for x in years_to_process])}]"

# PL projects id
if pl_projects_dict:
    candidate_id_columns = [
        "Номер БЦ",
        "Номер контракта",
        "Номер договора",
        "Номер проекта",
        "Project ID",
    ]

    prepared_projects_dict: dict[str, pd.DataFrame] = {}
    global_id_column: Optional[str] = None

    for key, df_month in pl_projects_dict.items():
        df_out = df_month.copy()
        month_id_column = next((col for col in candidate_id_columns if col in df_out.columns), None)

        if month_id_column is None:
            print(f"⚠️ Пропускаем {key}: не найден столбец идентификатора из {candidate_id_columns}")
            continue

        if global_id_column is None:
            global_id_column = month_id_column
        elif month_id_column != global_id_column:
            df_out = df_out.rename(columns={month_id_column: global_id_column})

        prepared_projects_dict[key] = df_out

    if prepared_projects_dict and global_id_column:
        report_path = report_common_name.replace("finance_report [", "finance_report_PL_projects [")
        result_path = write_monthly_with_highlights(
            dfs_dict=prepared_projects_dict,
            output_path=report_path,
            id_col=global_id_column,
            highlight_first=True,  # первый месяц считаем «появлениями»
            add_flag_column=True
        )
        print("Готово:", result_path)
    else:
        print("⚠️ Не удалось подготовить данные проектов для выгрузки.")

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
df_all = pd.concat(
    [row_with_profit_center_dict[k] for k in row_with_profit_center_dict],
    ignore_index=True
)
df_summary = df_pl_summary.merge(df_all, on=['Год', 'Месяц'], how='left')
save_summary_with_format(df_summary, report_path)
print(report_path)
print(f"Chargeability суммарный за год во вкладке `Chargeability_отчет` файла PL_XX {max(years_to_process)}.xlsx")
