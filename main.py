from pathlib import Path
import pandas as pd
from report_utils import load_pars, load_book, parse_raw, add_group_and_sort

def main():
    root = Path(__file__).parent
    final_dir = root / "final"
    final_dir.mkdir(exist_ok=True)

    pars_cols = load_pars(root)
    book = load_book(root)

    # Raw
    raw = parse_raw(root, pars_cols)

    # Grouped
    grouped = add_group_and_sort(raw.copy(), book)

    out_file = final_dir / "final_report.xlsx"
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="Raw", index=False)
        grouped.to_excel(writer, sheet_name="Grouped", index=False)

    print("✅ Отчёт сохранён в", out_file)

if __name__ == "__main__":
    main()
