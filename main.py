from pathlib import Path
import pandas as pd
from report_utils import (
    load_pars, load_book,
    parse_raw, prepare_grouped, prepare_logistics
)


def main():
    root = Path(__file__).parent
    final_dir = root / "final"
    final_dir.mkdir(exist_ok=True)

    pars_cols = load_pars(root)
    book = load_book(root)

    # Оригинальный отчёт
    pre_files = sorted((root / "pre").glob("*.xlsx"))
    if not pre_files:
        raise FileNotFoundError("Нет файлов в папке pre/")
    src = pre_files[0]
    original = pd.read_excel(src)

    # Short
    short = parse_raw(src, pars_cols)

    #checkpoint
    print("Колонки в short:", list(short.columns))

    # Grouped
    grouped = prepare_grouped(short, book)

    # Логистика
    logistics = prepare_logistics(original, book)

    # Запись в Excel
    out_file = final_dir / "final_report.xlsx"
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        original.to_excel(writer, sheet_name="Оригинальный отчет", index=False)
        short.to_excel(writer, sheet_name="Short", index=False)
        grouped.to_excel(writer, sheet_name="Grouped", index=False)
        logistics.to_excel(writer, sheet_name="Логистика", index=False)

    print("✅ Отчёт сохранён в", out_file)


if __name__ == "__main__":
    main()
