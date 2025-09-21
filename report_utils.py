import pandas as pd
import yaml
from pathlib import Path


# ---------- Загрузка конфигов ----------

def load_pars(root: Path) -> list[str]:
    """Загружает список колонок из pars.yaml"""
    pars_file = root / "pars.yaml"
    with open(pars_file, "r", encoding="utf-8") as f:
        pars_cols = yaml.safe_load(f)
    return pars_cols


def load_book(root: Path) -> dict:
    """Загружает словарь артикулов -> {group, color} из book.yaml"""
    book_file = root / "book.yaml"
    with open(book_file, "r", encoding="utf-8") as f:
        book = yaml.safe_load(f)
    return book


# ---------- Парсинг исходного отчёта ----------

def parse_raw(src: Path, pars_cols: list[str]) -> pd.DataFrame:
    """Парсит оригинальный Excel, оставляя только нужные колонки"""
    df = pd.read_excel(src)
    df.columns = df.columns.str.strip()
    cols = [c for c in pars_cols if c in df.columns]
    return df[cols].copy()


# ---------- Вспомогательные функции ----------

def normalize_key(s: str) -> str:
    """Нормализует ключ для поиска в book.yaml"""
    return str(s).strip().lower()


def add_group_and_color(df: pd.DataFrame, book: dict) -> pd.DataFrame:
    """Добавляет колонки Group и Color на основе book.yaml"""

    def lookup(key):
        key_norm = normalize_key(key)
        if key_norm in book:
            item = book[key_norm]
            if isinstance(item, dict):
                return item.get("group", "Unknown"), item.get("color", "Unknown")
            else:
                return item, "Unknown"  # старый формат
        return "Unknown", "Unknown"

    groups, colors = zip(*df["Артикул поставщика"].map(lookup))
    df.insert(0, "Color", colors)
    df.insert(0, "Group", groups)
    return df


def sort_by_book(df: pd.DataFrame, book: dict, extra_sort: list[str] | None = None) -> pd.DataFrame:
    """Сортировка по Group (порядок из book.yaml) + доп. колонки"""
    group_order = []
    for v in book.values():
        g = v["group"] if isinstance(v, dict) else v
        if g not in group_order:
            group_order.append(g)
    group_order.append("Unknown")

    df["Group_sort"] = pd.Categorical(df["Group"], categories=group_order, ordered=True)

    sort_cols = ["Group_sort", "Артикул поставщика"]
    if extra_sort:
        sort_cols.extend(extra_sort)

    df.sort_values(sort_cols, inplace=True, kind="mergesort")
    df.drop(columns="Group_sort", inplace=True)
    return df


# ---------- Листы финального отчёта ----------

def prepare_grouped(df: pd.DataFrame, book: dict) -> pd.DataFrame:
    """Готовит лист Grouped"""
    df = add_group_and_color(df.copy(), book)

    if "Обоснование для оплаты" in df.columns:
        pref = ["Продажа", "Логистика"]
        unique = [v for v in df["Обоснование для оплаты"].dropna().unique().tolist()]
        ob_order = [p for p in pref if p in unique] + sorted([v for v in unique if v not in pref])
        df["Ob_sort"] = pd.Categorical(df["Обоснование для оплаты"], categories=ob_order, ordered=True)
        df = sort_by_book(df, book, extra_sort=["Ob_sort"])
        df.drop(columns="Ob_sort", inplace=True)
    else:
        df = sort_by_book(df, book)

    return df


def prepare_logistics(df: pd.DataFrame, book: dict) -> pd.DataFrame:
    """Готовит лист Логистика"""
    if "Обоснование для оплаты" not in df.columns:
        return pd.DataFrame()

    logistics = df[df["Обоснование для оплаты"] == "Логистика"].copy()
    logistics = add_group_and_color(logistics, book)

    logistics_cols = [
        "Group",
        "Color",
        "Артикул поставщика",
        "Размер",
        "Обоснование для оплаты",
        "Услуги по доставке товара покупателю",
        "Дата заказа покупателем",
        "Дата продажи",
        "Виды логистики, штрафов и корректировок ВВ",
        "Количество доставок",
        "Количество возврата",
        "Склад",
        "Наименование офиса доставки",
        "Srid",
        "Код маркировки",
        "Страна",
        "Фиксированный коэффициент склада по поставке",
        "Дата начала действия фиксации",
        "Дата конца действия фиксации",
    ]
    logistics = logistics[[c for c in logistics_cols if c in logistics.columns]]

    # сортировка: Group → Color → Размер → Дата продажи
    logistics = sort_by_book(logistics, book, extra_sort=["Color", "Размер", "Дата продажи"])

    # добавляем строки ИТОГО
    if "Услуги по доставке товара покупателю" in logistics.columns:
        blocks = []
        for (g, c, size), block in logistics.groupby(["Group", "Color", "Размер"], dropna=False, sort=False):
            blocks.append(block)

            total_value = block["Услуги по доставке товара покупателю"].sum()
            total_row = {col: None for col in logistics.columns}
            total_row["Group"] = g
            total_row["Color"] = c
            total_row["Размер"] = size
            total_row["Артикул поставщика"] = "ИТОГО"
            total_row["Обоснование для оплаты"] = "Сумма"
            total_row["Услуги по доставке товара покупателю"] = total_value

            blocks.append(pd.DataFrame([total_row]))

        logistics = pd.concat(blocks, ignore_index=True)

    return logistics
