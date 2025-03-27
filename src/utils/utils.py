
from pathlib import Path
import pandas as pd


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

        return {
            "Файл": file_path.name,
            "Год": file_path.parent.name,
            "Месяц": file_path.stem.split('_')[1],
            "Profit Center": target_profit_center,
            "Сумма 'Реализация без НДС'": filtered_df["Реализация без НДС"].sum(),
            "Сумма 'Total Direct Costs'": filtered_df["Total Direct     Costs"].sum(),
            "Сумма 'Total Operating Costs'": filtered_df["Total Operating Costs"].sum()
        }
    except Exception as e:
        return {
            "Файл": file_path.name,
            "Год": file_path.parent.name,
            "Месяц": file_path.stem.split('_')[1],
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
