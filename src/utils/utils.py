import re
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def extract_profit_center_data(file_path: Path, target_profit_center: str):
    try:
        df = pd.read_excel(file_path, sheet_name="pl_projects")

        # Отбрасываем строки без Profit Center
        df_cleaned = df[df["Profit Center"].notna()]

        # Фильтрация по заданному Profit Center и непустой колонке "Реализация без НДС"
        filtered_df = df_cleaned[
            (df_cleaned["Profit Center"] == target_profit_center) &
            (df_cleaned["Реализация без НДС"].notna())
        ]

        # Извлекаем месяц из названия файла: PL_02 2025.xlsx -> 02
        match = re.search(r"PL_(\d{2})", file_path.stem)
        month = int(match.group(1)) if match else None
        year = int(file_path.parent.name)

        total_revenue = filtered_df["Реализация без НДС"].sum()
        total_direct_cost = filtered_df["Total Direct     Costs"].sum()
        total_operating_cost = filtered_df["Total Operating Costs"].sum()

        return {
            "Файл": file_path.name,
            "Год": year,
            "Месяц": month,
            "Profit Center": target_profit_center,
            "Сумма 'Реализация без НДС'": total_revenue,
            "Сумма 'Total Direct Costs'": total_direct_cost,
            "Сумма 'Total Operating Costs'": total_operating_cost,
            "Operation Profit": total_revenue-total_direct_cost-total_operating_cost
        }
    except Exception as e:
        return {
            "Файл": file_path.name,
            "Год": file_path.parent.name,
            "Месяц": None,
            "Ошибка": str(e)
        }


def process_all_pl_files(base_dir: str, years: list, target_profit_center: str):
    results = []

    for year in years:
        year_path = Path(base_dir) / str(year)
        if not year_path.exists():
            continue

        for file in sorted(year_path.glob("PL_*.xlsx")):
            print(f"Working with {file}")
            result = extract_profit_center_data(file, target_profit_center)
            results.append(result)

    return pd.DataFrame(results)


def save_summary_with_format(df, output_path):
    # Сохраняем без форматирования сначала
    df.to_excel(output_path, index=False)

    # Загружаем файл через openpyxl
    wb = load_workbook(output_path)
    ws = wb.active

    # Поиск нужных колонок
    headers = [cell.value for cell in ws[1]]
    money_columns = [
        "Сумма 'Реализация без НДС'",
        "Сумма 'Total Direct Costs'",
        "Сумма 'Total Operating Costs'",
        "Operation Profit"
    ]

    # Финансовый формат в рублях (с пробелами и запятыми)
    rub_format = '#,##0 ₽'  # Здесь между # и ##0 — не обычный пробел, а неразрывный (U+00A0)

    for col_name in money_columns:
        if col_name in headers:
            col_idx = headers.index(col_name) + 1  # Excel columns start at 1
            col_letter = ws.cell(row=1, column=col_idx).column_letter

            # Устанавливаем ширину столбца
            ws.column_dimensions[col_letter].width = 18

            # Применяем формат к значениям
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    cell.number_format = rub_format  # Финансовый формат

    wb.save(output_path)
