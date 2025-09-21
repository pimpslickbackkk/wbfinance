from pathlib import Path
import pandas as pd
from report_utils import load_pars, load_book, parse_raw, add_group_and_sort

def main():
    root = Path(__file__).parent
    final_dir = root / "final"
    final_dir.mkdir(exist_ok=True)

    pars_cols = load_pars(root)
    book = load_book(root)

    # Оригинальный отчёт (все данные как есть)
    pre_files = sorted((root / "pre").glob("*.xlsx"))
    if not pre_files:
        raise FileNotFoundError("Нет файлов в папке pre/")
    src = pre_files[0]
    original = pd.read_excel(src)

    # Short (только нужные колонки из pars.yaml)
    short = original[[c for c in pars_cols if c in original.columns]].copy()

    # Grouped (с группами и сортировкой)
    grouped = add_group_and_sort(short.copy(), book)

    out_file = final_dir / "final_report.xlsx"
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        original.to_excel(writer, sheet_name="Оригинальный отчет", index=False)
        short.to_excel(writer, sheet_name="Short", index=False)
        grouped.to_excel(writer, sheet_name="Grouped", index=False)

    print("✅ Отчёт сохранён в", out_file)


if __name__ == "__main__":
    main()
