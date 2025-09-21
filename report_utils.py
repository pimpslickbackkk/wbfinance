import pandas as pd
import yaml
from pathlib import Path

def load_yaml(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def load_pars(root: Path):
    data = load_yaml(root / "pars.yaml")
    return data["columns"]

def load_book(root: Path):
    data = load_yaml(root / "book.yaml")
    return data

def normalize_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = s.replace(" ", "").replace("-", "_").replace("/", "_")
    return s

def parse_raw(root: Path, pars_cols):
    pre_files = sorted((root / "pre").glob("*.xlsx"))
    if not pre_files:
        raise FileNotFoundError("Нет файлов в папке pre/")
    src = pre_files[0]
    df = pd.read_excel(src)

    # оставляем только нужные колонки
    cols = [c for c in pars_cols if c in df.columns]
    return df[cols].copy()

def add_group_and_sort(df: pd.DataFrame, book: dict):
    norm_map = {normalize_key(k): v for k, v in book.items()}
    df.insert(0, "Group", df["Артикул поставщика"].map(
        lambda x: norm_map.get(normalize_key(x), "Unknown"))
    )
    # сортировка: по Group (в порядке book.yaml), по Артикулу, по Обоснованию
    group_order = list(dict.fromkeys(book.values())) + ["Unknown"]
    df["Group_sort"] = pd.Categorical(df["Group"], categories=group_order, ordered=True)

    if "Обоснование для оплаты" in df.columns:
        pref = ["Продажа", "Логистика"]
        unique = [v for v in df["Обоснование для оплаты"].dropna().unique().tolist()]
        ob_order = [p for p in pref if p in unique] + sorted([v for v in unique if v not in pref])
        df["Ob_sort"] = pd.Categorical(df["Обоснование для оплаты"], categories=ob_order, ordered=True)
        df.sort_values(["Group_sort", "Артикул поставщика", "Ob_sort"], inplace=True, kind="mergesort")
        df.drop(columns="Ob_sort", inplace=True)
    else:
        df.sort_values(["Group_sort", "Артикул поставщика"], inplace=True)

    df.drop(columns="Group_sort", inplace=True)
    return df
