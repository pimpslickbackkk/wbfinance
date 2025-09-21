import pandas as pd
import yaml
from pathlib import Path


def load_pars(root: Path) -> list[str]:
    """Загружает список колонок из pars.yaml"""
    pars_file = root / "pars.yaml"
    with open(pars_file, "r", encoding="utf-8") as f:
        pars_cols = yaml.safe_load(f)
    return pars_cols


def load_book(root: Path) -> dict:
    """Загружает словарь групп из book.yaml"""
    book_file = root / "book.yaml"
    with open(book_file, "r", encoding="utf-8") as f:
        book = yaml.safe_load(f)
    return book


def parse_raw(src: Path, pars_cols: list[str]) -> pd.DataFrame:
    """Парсит оригинальный Excel, оставляя только нужные колонки"""
    df = pd.read_excel(src)
    cols = [c for c in pars_cols if c in df.columns]
    return df[cols].copy()


def normalize_key(s: str) -> str:
    """Нормализует ключ для поиска в book.yaml"""
    return str(s).strip().lower()


def add_group_column(df: pd.DataFrame, book: dict) -> pd.DataFrame:
    """Добавляет колонку Group на основе book.yaml"""
    norm_map = {normalize_key(k): v for k, v in book.items()}
    df.insert(
        0,
        "Group",
        df["Артикул поставщика"].map(lambda x: norm_map.get(normalize_key(x), "Unknown"))
    )
    return df


def sort_by_book(df: pd.DataFrame, book: dict, extra_sort: list[str] | None = None) -> pd.DataFrame:
    """
    Сортирует DataFrame по Group (порядок из book.yaml) + дополнительные колонки.
    """
    group_order = list(dict.fromkeys(book.values())) + ["Unknown"]
    df["Group_sort"] = pd.Categorical(df["Group"], categories=group_order, ordered=True)

    sort_cols = ["Group_sort", "Артикул поставщика"]
    if extra_sort:
        sort_cols.extend(extra_sort)

    df.sort_values(sort_cols, inplace=True, kind="mergesort")
    df.drop(columns="Group_sort", inplace=True)
    return df


def prepare_grouped(df: pd.DataFrame, book: dict) -> pd.DataFrame:
    """Готовит лист Grouped"""
    df = add_group_column(df.copy(), book)

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
    logistics = add_group_column(logistics, book)

    logistics_cols = [
        "Group",
        "Артикул поставщика",
        "Размер",
        "Обоснование для оплаты",
        "Srid",
        "Дата заказа покупателем",
        "Дата продажи",
        "Виды логистики, штрафов и корректировок ВВ",
        "Количество доставок",
        "Количество возврата",
        "Услуги по доставке товара покупателю",
        "Склад",
        "Наименование офиса доставки",
        "Код маркировки",
        "Страна",
        "Фиксированный коэффициент склада по поставке",
        "Дата начала действия фиксации",
        "Дата конца действия фиксации",
    ]
    logistics = logistics[[c for c in logistics_cols if c in logistics.columns]]

    logistics = sort_by_book(logistics, book, extra_sort=["Размер", "Srid"])
    return logistics
