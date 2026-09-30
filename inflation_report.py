# %%
"""
Месячный мониторинг инфляции.

Скрипт читает исходный Excel-файл, считает показатели в Python и сохраняет отчет
в Excel: на каждом листе таблица с данными (за весь период, с 2002 г.) и рядом
редактируемые Excel-графики (с 2024 г.), оформленные по брендбуку.
По желанию дополнительно формирует PDF с теми же графиками.

Листы отчета:
    1. Хедлайн и базовая — м/м (SA и nSA), г/г и SAAR хедлайна и четырех мер
       базовой инфляции;
    2. Вклады — разложение SA-хедлайна м/м на вклады (набор компонентов задается
       в настройках), месяцы на графике сгруппированы по кварталам;
    3. Накопленная — уровень цен в % к январю текущего года;
    4. Монетарная и немонетарная инфляция;
    5. Матрица «инфляция — разброс изменений цен по корзине»;
    6. Топ-5 категорий;
    7+. Отдельные категории (по названию столбца), если они заданы.

Запуск в Jupyter:  %run inflation_report.py
Все настройки собраны в блоке «НАСТРОЙКИ» сразу ниже. Зависимости — requirements.txt.
"""

from __future__ import annotations

from pathlib import Path

# ╔═════════════════════════════════════════════════════════════════════════════╗
# ║                                 НАСТРОЙКИ                                   ║
# ║            Всё, что выбирается перед запуском, — в разделах 1–9.            ║
# ╚═════════════════════════════════════════════════════════════════════════════╝

# ── 1. Файлы ──────────────────────────────────────────────────────────────────
INPUT_FILE = Path("Месячная инфляция_Данные.xlsx")        # исходные данные (рядом с кодом)
OUTPUT_XLSX = INPUT_FILE.with_name("Месячная инфляция_Отчет.xlsx")
EXPORT_PDF = False                                         # True — дополнительно сохранить PDF
OUTPUT_PDF = INPUT_FILE.with_name("Месячная инфляция_Графики.pdf")

# ── 2. С какого месяца строить графики ────────────────────────────────────────
# Таблицы на листах всегда выгружаются за весь период (с 2002 г.), графики —
# начиная с этого месяца. Чтобы продлить график назад, достаточно в Excel
# поменять первую строку диапазона ряда.
CHART_START = "2024-01"

# ── 3. Лист 1: хедлайн и базовая инфляция ─────────────────────────────────────
# На каких данных считать базовую инфляцию:
#   "SA"  — 44 крупные категории с листа «SA крупные категории». И отбор категорий
#           (волатильность, усечение), и агрегирование — по SA-данным.
#           История с 2002 г. Рекомендуемый вариант.
#   "nSA" — детальные категории (группа 0) без сезонной корректировки.
#           История только с 2024 г.; SAAR базовой инфляции не считается.
CORE_BASIS = "SA"

# Метод 1 — без 20% наиболее волатильных категорий.
CORE_VOLATILITY_WINDOW = 3      # окно в месяцах (включая текущий) для дисперсии м/м
CORE_VOLATILITY_SHARE = 0.20    # доля исключаемых категорий (по числу категорий)

# Метод 2 — усечение (trimmed mean): доли веса, отрезаемые снизу и сверху.
CORE_TRIM_LOW = 0.10
CORE_TRIM_HIGH = 0.10

# Метод 3 — исключение: исключаются категории, у которых есть любой из этих кодов.
# ЖКУ указаны сознательно: у крупных SA-категорий кода РУ нет, и регулируемые
# услуги там представлены только ЖКУ. У детальных категорий все ЖКУ и так
# помечены РУ, поэтому в варианте "nSA" результат от этого не меняется.
CORE_EXCLUDE_CODES = ["РУ", "ЖКУ", "ПО", "С", "А", "ТМ", "ТНМ"]
# Метод 4 — исключение без туризма: к кодам метода 3 добавляются эти.
CORE_EXCLUDE_TOURISM_CODES = ["ВТ", "ЗТ"]

# ── 4. Лист 2: вклады в хедлайн м/м ───────────────────────────────────────────
# "SA"  — вклады в SA-хедлайн по 44 крупным SA-категориям (история с 2002 г.).
# "nSA" — вклады в хедлайн без сезонной корректировки по детальным категориям
#         (с 2024 г.). Нужен, когда разложению нужны коды, которых нет у крупных
#         категорий (например, РУ — регулируемые услуги).
CONTRIBUTIONS_BASIS = "SA"

# Компоненты разложения: ("Подпись на графике", {условия}).
# Условия внутри одного компонента выполняются одновременно:
#   "код":       ["ПР"]      — есть ХОТЯ БЫ ОДИН из кодов (строка «Код», части через «_»)
#   "не код":    ["ПО"]      — нет НИ ОДНОГО из кодов
#   "Вал.курс":  1 или 0     — флаг из строки с таким названием на листе «Данные»;
#                              так же работают "Монетар. Инфляция", "Акцизы" и др.
#   "категории": ["Мясопродукты", ...] — конкретные категории по названию столбца
#   "остальное": True        — всё, что не попало в другие компоненты
# Категория попадает в ПЕРВЫЙ подходящий компонент сверху вниз, поэтому
# «остальные продовольственные» достаточно записать после «плодоовощей».
CONTRIBUTIONS = [
    ("Продовольственные товары", {"код": ["ПР"]}),
    ("Непродовольственные товары", {"код": ["НЕПР"]}),
    ("Услуги", {"код": ["У"]}),
]
# Примеры других разложений — раскомментируйте нужное:
# CONTRIBUTIONS = [
#     ("Плодоовощи", {"код": ["ПО"]}),
#     ("Остальные продовольственные", {"код": ["ПР"]}),
#     ("Непродовольственные товары", {"код": ["НЕПР"]}),
#     ("Услуги", {"код": ["У"]}),
# ]
# CONTRIBUTIONS = [
#     ("Подверженные колебаниям валютного курса", {"Вал.курс": 1}),
#     ("Не подверженные колебаниям валютного курса", {"Вал.курс": 0}),
# ]
# CONTRIBUTIONS = [                        # только с CONTRIBUTIONS_BASIS = "nSA"
#     ("Продовольственные товары", {"код": ["ПР"]}),
#     ("Непродовольственные товары", {"код": ["НЕПР"]}),
#     ("Регулируемые услуги", {"код": ["РУ"]}),
#     ("Нерегулируемые услуги", {"код": ["У"]}),
# ]

# ── 5. Лист 3: накопленная инфляция ───────────────────────────────────────────
# Уровень цен в % к базовому месяцу текущего года (в базовом месяце = 0):
# правее базы — рост цен (> 0), левее — более низкий уровень цен (< 0).
CUMULATIVE_BASE_MONTH = 1       # 1 = январь текущего года; 0 = декабрь прошлого года
# Хедлайн выводится всегда. Дополнительные линии — названия столбцов листа
# «Данные» или рассчитанных рядов ("Хедлайн SA", "Базовая: усечение 10%/10%",
# "Монетарная" и т.п. — полный список печатается при запуске).
CUMULATIVE_EXTRA_SERIES = []    # например: ["Продовольственные товары", "Услуги"]

# ── 6. Лист 4: монетарная и немонетарная инфляция ─────────────────────────────
# Детальные категории делятся по флагу: 1 → первая линия, 0 → вторая.
SPLIT_FLAG = "Монетар. Инфляция"
SPLIT_NAMES = ("Монетарная", "Немонетарная")
SPLIT_CHART = "г/г"             # что показывать на графике: "г/г" или "м/м"

# ── 7. Лист 5: матрица «инфляция — разброс по корзине» ────────────────────────
MATRIX_BASIS = "SA"             # "SA":  SA-хедлайн м/м и разброс SA м/м по 44 крупным категориям
                                # "nSA": хедлайн м/м и разброс м/м по детальным категориям
MATRIX_DISPERSION = "std"       # "std" — взвешенное стандартное отклонение м/м по корзине;
                                # "iqr" — взвешенный межквартильный размах (устойчив к выбросам;
                                #         для "nSA" лучше он: иначе разброс задают единичные позиции)
MATRIX_YEARS = 3                # сколько календарных лет наносить (включая текущий)
MATRIX_TARGET_SAAR = 4.0        # граница «высокая/низкая»: м/м, соответствующий 4% в год

# ── 8. Лист 6: топ-5 ──────────────────────────────────────────────────────────
TOP_GROUPS = ["1", "11", "111", "0"]
TOP_N = 5

# ── 9. Листы 7+: отдельные категории ──────────────────────────────────────────
# Названия столбцов листа «Данные» — любые, в том числе самые мелкие.
# Для каждой категории добавляется свой лист. Пустой список [] — листов нет.
SELECTED_CATEGORIES = ["Мясопродукты"]

# ── Технические настройки (менять обычно не нужно) ────────────────────────────
DATA_SHEET = "Данные"
SA_SHEET = "SA крупные категории"
CODE_IMPLIES = {"ЖКУ": ["У"]}           # ЖКУ — это услуги: код ЖКУ означает и У
VALIDATION_TOLERANCE_PP = 0.15          # допуск сверки с Росстатом, п.п. м/м
MIN_DETAILED_COVERAGE_WEIGHT = 95.0     # месяцы с меньшим покрытием веса детальными
                                        # категориями не считаются (в % веса)
COLUMN_ALIASES = {
    "headline": ["Все товары и услуги"],
    "headline_sa": ["SA Все товары и услуги"],
    "food": ["Продовольственные товары"],
    "nonfood": ["Непродовольственные товары"],
    "services": ["Услуги"],
}

# ╔═════════════════════════════════════════════════════════════════════════════╗
# ║                  ОФОРМЛЕНИЕ ПО БРЕНДБУКУ (версия 3.0)                       ║
# ╚═════════════════════════════════════════════════════════════════════════════╝
# Цвета — раздел 1.9 «Цвета для графиков и диаграмм». Шрифты — раздел 1.10:
# в офисных документах заголовки набираются Arial, основной текст — Times New
# Roman. Полужирное начертание запрещено, заголовки и подписи к графикам —
# фирменным серым (раздел 1.11). На графиках брендбука нет линий сетки.
BRAND = {
    "navy": "#1E3B56", "magenta": "#B1046E", "beige": "#A99892", "blue": "#009AD9",
    "coral": "#ED695A", "purple": "#6758A2", "green": "#6EBC84", "mustard": "#D5AD00",
    "grey_blue": "#818DA2", "pink": "#D492B6", "light_beige": "#D4CAC7",
    "light_blue": "#A2CCEE", "light_coral": "#F7BAAA", "light_purple": "#B1A8D3",
    "light_green": "#C0DFC3", "light_mustard": "#EBD592",
    "grey": "#A8B6BF", "light_grey": "#D1DADF",
    "text_grey": "#586C76",     # заголовки, подписи к графикам
    "text": "#000000",
}
# Порядок добавления цветов на «простых» графиках (брендбук, 1.9).
PALETTE_SIMPLE = [BRAND[k] for k in (
    "navy", "magenta", "beige", "blue", "grey_blue", "pink", "light_beige", "light_blue",
    "coral", "purple", "green", "mustard", "light_coral", "light_purple", "light_green",
    "light_mustard",
)]
FONT_TITLE = "Arial"
FONT_TEXT = "Times New Roman"
CHART_BACKGROUND = "#FFFFFF"
CHART_GRIDLINES = False
CHART_SIZE = (880, 430)                 # ширина и высота графиков в пикселях
DATE_AXIS_FORMAT = "[$-419]mmm yy"      # «янв 24» при любом языке Excel

# Базовая инфляция на графиках: оттенки розового, отличаются типом линии/маркера.
CORE_STYLES = {
    "vol": (BRAND["magenta"], "solid", "circle"),
    "trim": (BRAND["pink"], "solid", "square"),
    "ex1": (BRAND["magenta"], "dash", "diamond"),
    "ex2": (BRAND["pink"], "dash", "triangle"),
}
HEADLINE_COLOR = BRAND["navy"]
HEADLINE_NSA_COLOR = BRAND["grey"]

# =============================================================================
#                        ДАЛЕЕ — КОД (менять не нужно)
# =============================================================================

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Iterable
import difflib
import math
import re

import numpy as np
import pandas as pd


MONTHS_SHORT = {1: "янв", 2: "фев", 3: "мар", 4: "апр", 5: "май", 6: "июн",
                7: "июл", 8: "авг", 9: "сен", 10: "окт", 11: "ноя", 12: "дек"}
MONTHS_NOM = {1: "январь", 2: "февраль", 3: "март", 4: "апрель", 5: "май", 6: "июнь",
              7: "июль", 8: "август", 9: "сентябрь", 10: "октябрь", 11: "ноябрь",
              12: "декабрь"}
MONTHS_DAT = {1: "январю", 2: "февралю", 3: "марту", 4: "апрелю", 5: "маю", 6: "июню",
              7: "июлю", 8: "августу", 9: "сентябрю", 10: "октябрю", 11: "ноябрю",
              12: "декабрю"}
CORE_KEYS = ["vol", "trim", "ex1", "ex2"]


# =============================================================================
# 1. ЧТЕНИЕ И НОРМАЛИЗАЦИЯ ДАННЫХ
# =============================================================================


def _is_na(value) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def normalize_text(value) -> str:
    if _is_na(value):
        return ""
    text = str(value).replace("\xa0", " ").strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", text)


def flag_key(value) -> str:
    """Ключ строки разметки: без регистра, пробелов и точек ("Вал. курс" = "вал.курс")."""
    return re.sub(r"[\s.\-_]+", "", normalize_text(value))


def normalize_group(value) -> str | None:
    if _is_na(value):
        return None
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    text = str(value).strip().upper()
    return text or None


def normalize_code(value) -> str | None:
    if _is_na(value):
        return None
    text = str(value).replace("\xa0", " ").strip().upper()
    return text or None


def code_tokens(code: str | None) -> frozenset:
    """Составной код "ЖКУ_РУ" → {"ЖКУ", "РУ", "У"} (с учетом CODE_IMPLIES)."""
    if not isinstance(code, str) or not code:
        return frozenset()
    tokens = {t for t in re.split(r"[_\s]+", code.upper()) if t}
    implies = {k.upper(): {v.upper() for v in vals} for k, vals in CODE_IMPLIES.items()}
    changed = True
    while changed:
        changed = False
        for token, extra in implies.items():
            if token in tokens and not extra <= tokens:
                tokens |= extra
                changed = True
    return frozenset(tokens)


def parse_mixed_date(value) -> pd.Timestamp:
    """Читает и Excel-дату, и числовой serial date. Возвращает первое число месяца."""
    if _is_na(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value.year, value.month, 1)
    if isinstance(value, (int, float, np.integer, np.floating)):
        # Excel serial date. База 1899-12-30 учитывает особенности Excel.
        if float(value) > 20000:
            dt = pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
            return pd.Timestamp(dt.year, dt.month, 1)
        return pd.NaT
    parsed = pd.to_datetime(str(value), errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return pd.NaT
    return pd.Timestamp(parsed.year, parsed.month, 1)


def _excel_col_name(position: int) -> str:
    """0 → A, 1 → B, ..."""
    name = ""
    position += 1
    while position:
        position, rem = divmod(position - 1, 26)
        name = chr(65 + rem) + name
    return name


@dataclass
class InflationInput:
    values: pd.DataFrame        # индексы м/м (пред. месяц = 100): месяцы × категории
    meta: pd.DataFrame          # разметка категорий (строки над весами)
    flags: dict                 # ключ флага → название строки на листе «Данные»
    weights: pd.DataFrame       # годовые веса, развернутые по месяцам
    weight_rows: dict           # название строки весов → веса
    sa_values: pd.DataFrame     # SA-индексы крупных категорий (те же id столбцов)
    cols: dict                  # id ключевых столбцов (хедлайн и т.д.)
    latest: pd.Timestamp        # последний месяц с данными по хедлайну


def load_input(path: Path) -> InflationInput:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл с данными: {path.resolve()}")
    with pd.ExcelFile(path, engine="openpyxl") as xls:
        if DATA_SHEET not in xls.sheet_names:
            raise KeyError(f'В файле нет листа "{DATA_SHEET}". Листы: {xls.sheet_names}')
        raw = xls.parse(DATA_SHEET, header=None)
        sa_raw = xls.parse(SA_SHEET, header=None) if SA_SHEET in xls.sheet_names else None

    # Категории — непустые заголовки первой строки, начиная с колонки B.
    cat_pos = [j for j in range(1, raw.shape[1]) if normalize_text(raw.iat[0, j])]
    if not cat_pos:
        raise ValueError(f'На листе "{DATA_SHEET}" не найдены названия категорий.')
    ids = pd.RangeIndex(len(cat_pos), name="id")

    labels = raw.iloc[:, 0]
    date_rows, dates, weight_idx, meta_idx = [], [], [], []
    for r in range(1, raw.shape[0]):
        value = labels.iat[r]
        dt = parse_mixed_date(value)
        if not pd.isna(dt):
            date_rows.append(r)
            dates.append(dt)
        elif isinstance(value, str) and value.strip().startswith("Вес "):
            weight_idx.append(r)
        elif normalize_text(value) and not date_rows and not weight_idx:
            meta_idx.append(r)
    if not date_rows:
        raise ValueError("Не найдены строки с месячными датами.")

    # --- Разметка: все строки между заголовками и весами (Группа, Код, флаги).
    meta = pd.DataFrame(index=ids)
    meta["name"] = [str(raw.iat[0, j]).strip() for j in cat_pos]
    meta["name_norm"] = meta["name"].map(normalize_text)
    meta["excel_col"] = [_excel_col_name(j) for j in cat_pos]
    flags: dict[str, str] = {}
    for r in meta_idx:
        label = str(labels.iat[r]).strip()
        key = flag_key(label)
        row_values = [raw.iat[r, j] for j in cat_pos]
        if key == "группа":
            meta["group"] = [normalize_group(v) for v in row_values]
        elif key == "код":
            meta["code"] = [normalize_code(v) for v in row_values]
        else:
            meta[key] = pd.to_numeric(pd.Series(row_values, dtype="object"),
                                      errors="coerce").to_numpy(dtype=float)
            if key not in ("кодросстат", "последнийгод"):
                flags[key] = label
    for key, label in [("group", "Группа"), ("code", "Код")]:
        if key not in meta.columns:
            raise KeyError(f'На листе "{DATA_SHEET}" не найдена строка "{label}".')
    meta["tokens"] = meta["code"].map(code_tokens)

    # --- Годовые веса: все строки "Вес ...". "Вес 2027" добавится без правок кода.
    weight_rows: dict[str, np.ndarray] = {}
    for r in weight_idx:
        label = str(labels.iat[r]).strip()
        weight_rows[label] = pd.to_numeric(
            pd.Series([raw.iat[r, j] for j in cat_pos], dtype="object"), errors="coerce"
        ).to_numpy(dtype=float)

    # --- Индексы.
    values = raw.iloc[date_rows, cat_pos].apply(pd.to_numeric, errors="coerce")
    values.index = pd.DatetimeIndex(dates, name="Дата")
    values.columns = ids
    values = values.sort_index()
    if values.index.duplicated().any():
        dup = values.index[values.index.duplicated()].strftime("%Y-%m").tolist()
        raise ValueError(f"В данных есть дублирующиеся месяцы: {dup}")
    # Полная месячная сетка: пропущенный месяц становится NaN, а не «склеивается».
    grid = pd.date_range(values.index.min(), values.index.max(), freq="MS", name="Дата")
    values = values.reindex(grid).astype(float)

    weights = pd.DataFrame(np.nan, index=grid, columns=ids)
    missing = []
    for year in sorted(set(grid.year)):
        label = f"Вес {year}"
        mask = grid.year == year
        if label in weight_rows:
            weights.loc[mask, :] = np.tile(weight_rows[label], (int(mask.sum()), 1))
        else:
            missing.append(label)
    if missing:
        print(f"⚠️ Не найдены строки весов: {', '.join(missing)}. Эти месяцы в расчетах будут пустыми.")

    cols = {key: find_column(meta, aliases, required=(key == "headline"))
            for key, aliases in COLUMN_ALIASES.items()}
    latest = values[cols["headline"]].last_valid_index()
    if latest is None:
        raise ValueError("В колонке хедлайна нет данных.")

    sa_values = _load_sa_sheet(sa_raw, meta, grid)
    return InflationInput(values=values, meta=meta, flags=flags, weights=weights,
                          weight_rows=weight_rows, sa_values=sa_values, cols=cols,
                          latest=latest)


def _load_sa_sheet(raw: pd.DataFrame | None, meta: pd.DataFrame, grid: pd.DatetimeIndex) -> pd.DataFrame:
    """SA-индексы крупных категорий. Столбцы сопоставляются с листом «Данные» по названию."""
    if raw is None:
        print(f'⚠️ В файле нет листа "{SA_SHEET}": расчеты на SA-данных невозможны.')
        return pd.DataFrame(index=grid, dtype=float)
    rows = [(r, parse_mixed_date(raw.iat[r, 0])) for r in range(1, raw.shape[0])]
    rows = [(r, d) for r, d in rows if not pd.isna(d)]
    sa_flagged = set(meta.index[meta["sa"].eq(1)]) if "sa" in meta.columns else None

    series, not_found = {}, []
    for j in range(1, raw.shape[1]):
        name = normalize_text(raw.iat[0, j])
        if not name:
            continue
        ids = meta.index[meta["name_norm"] == name].tolist()
        if sa_flagged is not None and len(ids) > 1:
            ids = [i for i in ids if i in sa_flagged] or ids
        if not ids:
            not_found.append(str(raw.iat[0, j]))
            continue
        s = pd.to_numeric(pd.Series([raw.iat[r, j] for r, _ in rows], dtype="object"),
                          errors="coerce")
        s.index = pd.DatetimeIndex([d for _, d in rows])
        series[ids[0]] = s[~s.index.duplicated()].reindex(grid)

    if not_found:
        print(f'⚠️ Столбцы листа "{SA_SHEET}" не найдены на листе "{DATA_SHEET}": {not_found}')
    if sa_flagged is not None:
        lost = sorted(sa_flagged - set(series))
        if lost:
            print(f'⚠️ Категории с SA = 1 отсутствуют на листе "{SA_SHEET}": '
                  f'{meta.loc[lost, "name"].tolist()}')
    return pd.DataFrame(series, index=grid, dtype=float)


def find_column(meta: pd.DataFrame, aliases: Iterable[str], *, required: bool = True) -> int | None:
    alias_norm = {normalize_text(x) for x in aliases}
    matches = meta.index[meta["name_norm"].isin(alias_norm)].tolist()
    if len(matches) == 1:
        return int(matches[0])
    if len(matches) > 1:
        raise ValueError(f"Найдено несколько колонок для {list(aliases)}: "
                         f"{meta.loc[matches, 'name'].tolist()}")
    if required:
        suggestions = difflib.get_close_matches(list(aliases)[0], meta["name"].tolist(), n=5, cutoff=0.5)
        raise KeyError(f"Не найдена колонка из вариантов {list(aliases)}. Похожие: {suggestions}")
    return None


def find_category(meta: pd.DataFrame, name: str) -> int | None:
    """Категория по названию столбца (без учета регистра и лишних пробелов)."""
    target = normalize_text(name)
    ids = meta.index[meta["name_norm"] == target].tolist()
    if len(ids) > 1:
        print(f"⚠️ Название «{name}» встречается несколько раз "
              f"(колонки {meta.loc[ids, 'excel_col'].tolist()}), взята первая.")
    if ids:
        return int(ids[0])
    contains = [n for n in meta["name"] if target and target in normalize_text(n)][:5]
    close = difflib.get_close_matches(name, meta["name"].tolist(), n=5, cutoff=0.5)
    suggestions = list(dict.fromkeys(contains + close))[:7]
    print(f"⚠️ Категория «{name}» не найдена на листе «{DATA_SHEET}». Похожие: {suggestions}")
    return None


# =============================================================================
# 2. ИНДЕКСНАЯ АРИФМЕТИКА
# =============================================================================


def yoy_from_mm(mm):
    """Г/г из цепочки м/м (%): произведение 12 последовательных месячных индексов."""
    log_factor = np.log1p(mm.astype(float) / 100.0)
    return np.expm1(log_factor.rolling(12, min_periods=12).sum()) * 100.0


def saar_from_mm(mm):
    """SAAR: м/м в годовом выражении."""
    return ((1.0 + mm / 100.0) ** 12 - 1.0) * 100.0


def saar_3m(mm):
    """SAAR 3м/3м: прирост за последние 3 месяца в годовом выражении."""
    log_factor = np.log1p(mm.astype(float) / 100.0)
    return np.expm1(log_factor.rolling(3, min_periods=3).sum() * 4.0) * 100.0


def laspeyres_weights(inp: InflationInput, cols) -> pd.DataFrame:
    """
    Веса для агрегирования м/м по схеме Росстата (цепной индекс Ласпейреса).

    Вес категории в месяце t = годовой вес × накопленный с декабря прошлого года
    индекс категории по месяц t-1 — так же, как на листе «Расчет прочих».
    В январе веса равны годовым. Для SA-агрегатов используются те же веса
    (по nSA-индексам), что и в столбце «SA Все товары и услуги».
    """
    cols = list(cols)
    ratio = (inp.values[cols] / 100.0).fillna(1.0)
    years = ratio.index.year
    cumulative = ratio.groupby(years).cumprod()
    cumulative_prev = cumulative.groupby(years).shift(1).fillna(1.0)
    return inp.weights[cols] * cumulative_prev


def weighted_mm(mm: pd.DataFrame, ew: pd.DataFrame) -> pd.Series:
    """Взвешенное среднее м/м по доступным в каждом месяце категориям."""
    w = ew.where(mm.notna() & (ew > 0))
    return (mm * w).sum(axis=1, min_count=1) / w.sum(axis=1, min_count=1)


def detailed_coverage(inp: InflationInput) -> pd.Series:
    """Доля веса (в %), покрытая детальными категориями (группа 0) с данными."""
    cols0 = inp.meta.index[inp.meta["group"].eq("0")]
    w = inp.weights[cols0].where(inp.values[cols0].notna())
    return w.sum(axis=1, min_count=1)


def level_from_mm(mm: pd.Series) -> pd.Series:
    """
    Уровень цен из цепочки м/м. За месяц до первого наблюдения уровень = 1.
    После пропуска внутри ряда уровень не определен (NaN).
    """
    s = mm.astype(float)
    first = s.first_valid_index()
    if first is None:
        return s * np.nan
    s = s.loc[first:]
    broken = s.isna().cumsum() > 0
    level = (1.0 + s.fillna(0.0) / 100.0).cumprod()
    level[broken] = np.nan
    before = first - pd.DateOffset(months=1)
    if before in mm.index:
        level = pd.concat([pd.Series([1.0], index=[before]), level])
    return level.reindex(mm.index)


def ytd_by_year(mm: pd.Series) -> pd.DataFrame:
    """Накопленная с начала года (с декабря прошлого года) инфляция: месяцы × годы."""
    out = {}
    for year in sorted(set(mm.dropna().index.year)):
        s = mm[mm.index.year == year]
        if s.empty or pd.isna(s.iloc[0]) or s.index[0].month != 1:
            continue  # без январских данных накопленную с начала года не посчитать
        broken = s.isna().cumsum() > 0
        path = ((1.0 + s / 100.0).cumprod() - 1.0) * 100.0
        path[broken] = np.nan
        out[year] = pd.Series(path.values, index=s.index.month)
    table = pd.DataFrame(out).reindex(range(1, 13))
    table.index.name = "Месяц"
    return table


def base_date(latest: pd.Timestamp) -> pd.Timestamp:
    month = int(CUMULATIVE_BASE_MONTH)
    if month == 0:
        return pd.Timestamp(latest.year - 1, 12, 1)
    if 1 <= month <= 12:
        return pd.Timestamp(latest.year, month, 1)
    raise ValueError("CUMULATIVE_BASE_MONTH: 0 (декабрь прошлого года) или 1–12.")


# =============================================================================
# 3. ВЫБОР КАТЕГОРИЙ ПО КОДАМ И ФЛАГАМ
# =============================================================================
# Флаги 0/1 нигде не используются как множители: категория выбирается явным
# сравнением «флаг = 1» или «флаг = 0». Пустой флаг не попадает ни туда, ни туда.

_KEY_CODE = {"код", "коды"}
_KEY_NOT_CODE = {"некод", "некоды"}
_KEY_NAMES = {"категории", "категория"}
_KEY_REST = {"остальное", "остальные"}


def _as_list(value) -> list:
    if isinstance(value, (list, tuple, set, frozenset)):
        return list(value)
    return [value]


def _has_any_code(tokens: pd.Series, codes: set) -> pd.Series:
    return tokens.map(lambda t: bool(t & codes)).astype(bool)


def select_categories(inp: InflationInput, spec: dict, candidates) -> pd.Series:
    meta = inp.meta.loc[list(candidates)]
    mask = pd.Series(True, index=meta.index)
    for raw_key, value in spec.items():
        key = flag_key(raw_key)
        if key in _KEY_REST:
            continue
        if key in _KEY_CODE:
            mask &= _has_any_code(meta["tokens"], {str(c).upper() for c in _as_list(value)})
        elif key in _KEY_NOT_CODE:
            mask &= ~_has_any_code(meta["tokens"], {str(c).upper() for c in _as_list(value)})
        elif key in _KEY_NAMES:
            names = {normalize_text(v) for v in _as_list(value)}
            unknown = names - set(inp.meta["name_norm"])
            if unknown:
                print(f"⚠️ Категории не найдены на листе «{DATA_SHEET}»: {sorted(unknown)}")
            mask &= meta["name_norm"].isin(names)
        elif key in inp.flags:
            if value not in (0, 1) or isinstance(value, bool):
                raise ValueError(f"Условие «{raw_key}»: флаг может быть только 1 или 0.")
            if value == 1:
                mask &= meta[key].eq(1)
            else:
                mask &= meta[key].eq(0)
        else:
            raise KeyError(
                f"Неизвестное условие «{raw_key}». Доступно: «код», «не код», «категории», "
                f"«остальное» и флаги: {list(inp.flags.values())}"
            )
    return mask


def assign_components(inp: InflationInput, components: list, candidates) -> tuple[pd.Series, list]:
    """
    Распределяет категории по компонентам (первый подходящий сверху вниз).
    Возвращает номер компонента для каждой категории (-1 — не распределена).
    """
    candidates = list(candidates)
    assign = pd.Series(-1, index=candidates)
    names, rest = [], None
    for k, (name, spec) in enumerate(components):
        names.append(name)
        keys = {flag_key(x) for x in spec}
        if keys & _KEY_REST:
            if rest is not None:
                raise ValueError("«остальное» можно указать только в одном компоненте.")
            rest = k
            continue
        mask = select_categories(inp, spec, candidates)
        taken = mask & (assign >= 0)
        if taken.any():
            print(f"ℹ️ Вклады: «{name}» — {int(taken.sum())} кат. уже учтены выше и не задваиваются.")
        assign[mask & (assign < 0)] = k
    if rest is not None:
        assign[assign < 0] = rest
    return assign, names


# =============================================================================
# 4. БАЗОВАЯ ИНФЛЯЦИЯ
# =============================================================================


@dataclass
class Basis:
    """Набор категорий, на котором считаются базовая инфляция, вклады, разброс."""
    kind: str                   # "SA" или "nSA"
    title: str                  # подпись для графиков и сообщений
    cols: list
    mm: pd.DataFrame            # м/м, %
    ew: pd.DataFrame            # веса Ласпейреса
    tokens: pd.Series
    names: pd.Series


def make_basis(inp: InflationInput, kind: str, coverage: pd.Series) -> Basis:
    kind_norm = str(kind).strip().lower()
    if kind_norm == "sa":
        cols = list(inp.sa_values.columns)
        if not cols:
            raise ValueError(f'Нет SA-данных (лист "{SA_SHEET}"): выберите вариант "nSA".')
        mm = inp.sa_values[cols] - 100.0
        title = f"крупные категории с сезонной корректировкой (SA, {len(cols)} шт.)"
    elif kind_norm == "nsa":
        cols = list(inp.meta.index[inp.meta["group"].eq("0")])
        mm = (inp.values[cols] - 100.0).copy()
        bad = ~(coverage >= MIN_DETAILED_COVERAGE_WEIGHT)
        mm.loc[bad.reindex(mm.index, fill_value=True).to_numpy()] = np.nan
        title = "детальные категории без сезонной корректировки (nSA)"
    else:
        raise ValueError(f'Неизвестный вариант данных "{kind}": нужен "SA" или "nSA".')
    return Basis(kind="SA" if kind_norm == "sa" else "nSA", title=title, cols=cols,
                 mm=mm, ew=laspeyres_weights(inp, cols),
                 tokens=inp.meta.loc[cols, "tokens"], names=inp.meta.loc[cols, "name"])


def core_labels() -> dict:
    return {
        "vol": f"без {CORE_VOLATILITY_SHARE * 100:.0f}% наиболее волатильных",
        "trim": f"усечение {CORE_TRIM_LOW * 100:.0f}%/{CORE_TRIM_HIGH * 100:.0f}%",
        "ex1": "исключение",
        "ex2": "исключение, без туризма",
    }


def weighted_trimmed_mean(x: np.ndarray, w: np.ndarray, low: float, high: float) -> tuple[float, np.ndarray]:
    """
    Взвешенное усеченное среднее: отрезаются доли low и high накопленного веса
    снизу и сверху распределения м/м. Категория на границе входит частично.
    Возвращает значение и эффективные веса (в исходном порядке).
    """
    order = np.argsort(x, kind="mergesort")
    share = w[order] / w.sum()
    upper = np.cumsum(share)
    lower = upper - share
    kept = np.clip(np.minimum(upper, 1.0 - high) - np.maximum(lower, low), 0.0, None)
    effective = np.empty_like(kept)
    effective[order] = kept
    if kept.sum() <= 0:
        return np.nan, effective
    return float((x[order] * kept).sum() / kept.sum()), effective


def core_measures(basis: Basis, latest: pd.Timestamp) -> tuple[pd.DataFrame, dict]:
    """
    Четыре меры базовой инфляции (м/м, %) и диагностика на последнюю дату.

    1) Без 20% наиболее волатильных: в каждом месяце t по окну из последних
       CORE_VOLATILITY_WINDOW месяцев (t включительно) считается дисперсия м/м
       каждой категории; исключаются 20% категорий с наибольшей дисперсией.
       Категории без полного окна (новые) не ранжируются и остаются в корзине.
    2) Усечение: взвешенное среднее без 10% веса снизу и 10% сверху.
    3) Исключение: без категорий с кодами CORE_EXCLUDE_CODES.
    4) Исключение без туризма: дополнительно без CORE_EXCLUDE_TOURISM_CODES.
    Во всех методах оставшаяся корзина перенормируется на 100%.
    """
    window = int(CORE_VOLATILITY_WINDOW)
    if window < 2:
        raise ValueError("CORE_VOLATILITY_WINDOW должно быть не меньше 2.")
    ex1 = {c.upper() for c in CORE_EXCLUDE_CODES}
    ex2 = ex1 | {c.upper() for c in CORE_EXCLUDE_TOURISM_CODES}
    keep1 = (~_has_any_code(basis.tokens, ex1)).to_numpy()
    keep2 = (~_has_any_code(basis.tokens, ex2)).to_numpy()

    values = basis.mm.to_numpy(dtype=float)
    weights = basis.ew.to_numpy(dtype=float)
    out = np.full((len(basis.mm), 4), np.nan)
    diag: dict = {}
    names = basis.names.to_numpy()
    p_latest = basis.mm.index.get_loc(latest) if latest in basis.mm.index else None

    for p in range(len(basis.mm)):
        x, w = values[p], weights[p]
        cur = ~np.isnan(x) & ~np.isnan(w) & (w > 0)
        if not cur.any():
            continue

        def wmean(mask):
            return float((x[mask] * w[mask]).sum() / w[mask].sum()) if mask.any() else np.nan

        # 1) Без наиболее волатильных.
        if p >= window - 1:
            idx = np.flatnonzero(cur)
            hist = values[p - window + 1:p + 1, idx]
            full = ~np.isnan(hist).any(axis=0)
            eligible = idx[full]
            variance = hist[:, full].var(axis=0, ddof=1)
            n_drop = math.ceil(round(CORE_VOLATILITY_SHARE * len(eligible), 9))
            dropped = eligible[np.argsort(-variance, kind="mergesort")[:n_drop]]
            keep_vol = cur.copy()
            keep_vol[dropped] = False
            out[p, 0] = wmean(keep_vol)
        # 2) Усечение.
        out[p, 1], effective = weighted_trimmed_mean(x[cur], w[cur], CORE_TRIM_LOW, CORE_TRIM_HIGH)
        # 3–4) Исключение.
        out[p, 2] = wmean(cur & keep1)
        out[p, 3] = wmean(cur & keep2)

        if p == p_latest:
            share = w[cur] / w[cur].sum()
            trimmed_share = 1.0 - effective / np.where(share > 0, share, np.nan)
            order = np.argsort(x[cur], kind="mergesort")
            median_x = x[cur][order][np.searchsorted(np.cumsum(share[order]), 0.5)]
            cur_names = names[cur]
            diag = {
                "vol_dropped": list(names[dropped]) if p >= window - 1 else [],
                "vol_eligible": int(len(eligible)) if p >= window - 1 else 0,
                "trim_low": sorted([(n, v) for n, v, t in zip(cur_names, x[cur], trimmed_share)
                                    if t > 1e-9 and v <= median_x], key=lambda item: item[1]),
                "trim_high": sorted([(n, v) for n, v, t in zip(cur_names, x[cur], trimmed_share)
                                     if t > 1e-9 and v > median_x], key=lambda item: -item[1]),
                "ex1": list(names[cur & ~keep1]),
                "ex2": list(names[cur & ~keep2]),
                "ex1_share": float(w[cur & ~keep1].sum() / w[cur].sum() * 100),
                "ex2_share": float(w[cur & ~keep2].sum() / w[cur].sum() * 100),
                "n": int(cur.sum()),
            }

    result = pd.DataFrame(out, index=basis.mm.index, columns=CORE_KEYS)
    missing_ex = sorted(c for c in ex2 if not _has_any_code(basis.tokens, {c}).any())
    diag["codes_not_found"] = missing_ex
    return result, diag


# =============================================================================
# 5. РАСЧЕТ ЛИСТОВ ОТЧЕТА
# =============================================================================


@dataclass
class ReportData:
    latest: pd.Timestamp
    chart_start: pd.Timestamp
    main: pd.DataFrame
    main_info: dict
    contrib: pd.DataFrame
    contrib_info: dict
    cumulative: pd.DataFrame
    cumulative_info: dict
    split: pd.DataFrame
    split_info: dict
    matrix: pd.DataFrame
    matrix_info: dict
    top_data: dict
    categories: list = field(default_factory=list)


def _cut(df: pd.DataFrame, latest: pd.Timestamp) -> pd.DataFrame:
    """Таблица до последнего месяца, без пустых строк в начале."""
    df = df.loc[:latest]
    first = df.dropna(how="all").index.min()
    return df.loc[first:] if first is not None and not pd.isna(first) else df.iloc[0:0]


def calc_main(inp, hl_mm, hl_sa_mm, core, basis) -> tuple[pd.DataFrame, dict]:
    labels = core_labels()
    tag = basis.kind
    table = {"Хедлайн nSA, м/м, %": hl_mm, "Хедлайн SA, м/м, %": hl_sa_mm}
    for k in CORE_KEYS:
        table[f"Базовая {tag}: {labels[k]}, м/м, %"] = core[k]
    table["Хедлайн, г/г, %"] = yoy_from_mm(hl_mm)
    for k in CORE_KEYS:
        table[f"Базовая: {labels[k]}, г/г, %"] = yoy_from_mm(core[k])
    table["Хедлайн SAAR, %"] = saar_from_mm(hl_sa_mm)
    if basis.kind == "SA":
        for k in CORE_KEYS:
            table[f"Базовая: {labels[k]}, SAAR, %"] = saar_from_mm(core[k])
    table["Хедлайн SAAR 3м/3м, %"] = saar_3m(hl_sa_mm)
    df = _cut(pd.DataFrame(table), inp.latest)
    df.insert(0, "Дата", df.index)
    return df.reset_index(drop=True), {"labels": labels, "kind": basis.kind, "title": basis.title}


def quarter_gapped(df: pd.DataFrame) -> pd.DataFrame:
    """Помесячная таблица → блоки по кварталам с пустой строкой между ними."""
    rows = []
    for (year, quarter), block in df.groupby([df.index.year, df.index.quarter]):
        for i, (dt, values) in enumerate(block.iterrows()):
            rows.append({"Квартал": f"{quarter} кв. {year}" if i == 0 else "",
                         "Месяц": MONTHS_SHORT[dt.month], "Дата": dt, **values.to_dict()})
        rows.append({"Квартал": " ", "Месяц": "", "Дата": pd.NaT})
    return pd.DataFrame(rows[:-1], columns=["Квартал", "Месяц", "Дата", *df.columns])


def calc_contributions(inp, basis, hl_mm, hl_sa_mm) -> tuple[pd.DataFrame, dict]:
    assign, names = assign_components(inp, CONTRIBUTIONS, basis.cols)
    unassigned = assign.index[assign < 0]
    if len(unassigned):
        w_last = basis.ew.loc[inp.latest, unassigned]
        share = float(w_last.sum() / basis.ew.loc[inp.latest].sum() * 100)
        print(f"⚠️ Вклады: {len(unassigned)} кат. ({share:.1f}% веса) не попали ни в один компонент "
              f"и показаны как «Не распределено»: {inp.meta.loc[unassigned[:8], 'name'].tolist()}"
              f"{' …' if len(unassigned) > 8 else ''}")
        assign[unassigned] = len(names)
        names = names + ["Не распределено"]

    mm, ew = basis.mm, basis.ew
    valid = mm.notna() & ew.notna() & (ew > 0)
    w = ew.where(valid)
    denominator = w.sum(axis=1, min_count=1)
    parts = {}
    for k, name in enumerate(names):
        cols_k = assign.index[assign == k]
        if len(cols_k) == 0:
            print(f"⚠️ Вклады: в компонент «{name}» не попало ни одной категории "
                  f"({basis.title}) — он не показывается.")
            continue
        parts[f"Вклад: {name}, п.п."] = (mm[cols_k] * w[cols_k]).sum(axis=1, min_count=1) / denominator

    if basis.kind == "SA":
        line_name, line = "Хедлайн SA, м/м, %", hl_sa_mm
    else:
        line_name, line = "Хедлайн nSA, м/м, %", hl_mm
    table = pd.DataFrame({line_name: line, **parts})
    residual = line - pd.DataFrame(parts).sum(axis=1, min_count=1)
    if residual.abs().max() > 0.005:
        table["Расхождение с хедлайном, п.п."] = residual
    table = _cut(table.loc[denominator.dropna().index.min():], inp.latest)
    info = {"line": line_name, "components": [c for c in table.columns if c != line_name],
            "basis": basis}
    return quarter_gapped(table), info


def resolve_series(inp: InflationInput, name: str, derived: dict) -> pd.Series | None:
    lookup = {normalize_text(k): k for k in derived}
    if normalize_text(name) in lookup:
        return derived[lookup[normalize_text(name)]]
    cid = find_category(inp.meta, name)
    return None if cid is None else inp.values[cid] - 100.0


def calc_cumulative(inp, hl_mm, derived) -> tuple[pd.DataFrame, dict]:
    base = base_date(inp.latest)
    base_label = f"{MONTHS_DAT[base.month]} {base.year} г."
    series = {"Хедлайн": hl_mm}
    for name in CUMULATIVE_EXTRA_SERIES:
        s = resolve_series(inp, name, derived)
        if s is not None:
            series[name] = s
    table = {}
    for name, mm in series.items():
        level = level_from_mm(mm)
        if base not in level.index or pd.isna(level.get(base)):
            print(f"⚠️ Накопленная: у ряда «{name}» нет данных за базовый месяц ({base:%m.%Y}).")
            continue
        table[f"{name}, % к {base_label}"] = (level / level.loc[base] - 1.0) * 100.0
    df = _cut(pd.DataFrame(table), inp.latest)
    df.insert(0, "Дата", df.index)
    return df.reset_index(drop=True), {"base": base, "base_label": base_label}


def calc_split(inp, coverage) -> tuple[pd.DataFrame, dict, dict]:
    key = flag_key(SPLIT_FLAG)
    if key not in inp.flags:
        raise KeyError(f"Флаг «{SPLIT_FLAG}» не найден. Доступны: {list(inp.flags.values())}")
    basis = make_basis(inp, "nSA", coverage)
    flag = inp.meta.loc[basis.cols, key]
    cols1 = flag.index[flag.eq(1)]
    cols0 = flag.index[flag.eq(0)]
    s1 = weighted_mm(basis.mm[cols1], basis.ew[cols1])
    s0 = weighted_mm(basis.mm[cols0], basis.ew[cols0])
    n1, n0 = SPLIT_NAMES
    table = pd.DataFrame({
        f"{n1}, м/м, %": s1, f"{n0}, м/м, %": s0,
        f"{n1}, г/г, %": yoy_from_mm(s1), f"{n0}, г/г, %": yoy_from_mm(s0),
    })
    df = _cut(table, inp.latest)
    df.insert(0, "Дата", df.index)

    # Контроль разметки: вес детальных категорий без флага на последнюю дату.
    active = basis.mm.loc[inp.latest].notna() & basis.ew.loc[inp.latest].gt(0)
    unflagged = active & ~(flag.eq(1) | flag.eq(0))
    if unflagged.any():
        share = basis.ew.loc[inp.latest, unflagged].sum() / basis.ew.loc[inp.latest, active].sum() * 100
        print(f"⚠️ {n1}/{n0.lower()}: у {int(unflagged.sum())} детальных категорий "
              f"({share:.2f}% веса) не заполнен флаг «{inp.flags[key]}».")
    return df.reset_index(drop=True), {"names": SPLIT_NAMES}, {n1: s1, n0: s0}


def weighted_dispersion(mm: pd.DataFrame, ew: pd.DataFrame, kind: str) -> pd.Series:
    valid = mm.notna() & ew.notna() & (ew > 0)
    w = ew.where(valid)
    w = w.div(w.sum(axis=1), axis=0)
    if kind == "std":
        mean = (mm * w).sum(axis=1, min_count=1)
        variance = (w * mm.sub(mean, axis=0) ** 2).sum(axis=1, min_count=1)
        return np.sqrt(variance)
    if kind == "iqr":
        out = pd.Series(np.nan, index=mm.index)
        for dt in mm.index:
            ok = valid.loc[dt].to_numpy()
            if not ok.any():
                continue
            x = mm.loc[dt].to_numpy()[ok]
            ww = w.loc[dt].to_numpy()[ok]
            order = np.argsort(x, kind="mergesort")
            cum = np.cumsum(ww[order])
            q = lambda p: x[order][min(np.searchsorted(cum, p), len(x) - 1)]
            out[dt] = q(0.75) - q(0.25)
        return out
    raise ValueError('MATRIX_DISPERSION: "std" или "iqr".')


def calc_matrix(inp, basis, hl_mm, hl_sa_mm) -> tuple[pd.DataFrame, dict]:
    kind = str(MATRIX_DISPERSION).strip().lower()
    disp = weighted_dispersion(basis.mm, basis.ew, kind)
    infl = hl_sa_mm if basis.kind == "SA" else hl_mm
    df = pd.DataFrame({"x": disp, "y": infl}).loc[:inp.latest].dropna()
    last3 = df.index[-3:]
    tier = np.where(df.index.isin(last3), 0, np.where(df.index.year == inp.latest.year, 1, 2))
    tier_names = ["Последние 3 месяца", f"Ранее в {inp.latest.year} г.", "Предыдущие годы"]
    disp_name = ("Разброс м/м по корзине (станд. откл.), п.п." if kind == "std"
                 else "Разброс м/м по корзине (межкварт. размах), п.п.")
    infl_name = "Хедлайн SA, м/м, %" if basis.kind == "SA" else "Хедлайн nSA, м/м, %"
    table = pd.DataFrame({
        "Дата": df.index,
        "Метка": [f"{MONTHS_SHORT[d.month]}.{d.year % 100:02d}" for d in df.index],
        disp_name: df["x"].to_numpy(),
        infl_name: df["y"].to_numpy(),
    })
    for t, name in enumerate(tier_names):
        table[name] = np.where(tier == t, df["y"].to_numpy(), np.nan)
    window_start = pd.Timestamp(inp.latest.year - int(MATRIX_YEARS) + 1, 1, 1)
    in_window = table["Дата"] >= window_start
    threshold = ((1.0 + MATRIX_TARGET_SAAR / 100.0) ** (1.0 / 12.0) - 1.0) * 100.0
    info = {"basis": basis, "tiers": tier_names, "x": disp_name, "y": infl_name,
            "window_start": window_start, "y_threshold": threshold,
            "x_threshold": float(table.loc[in_window, disp_name].median())}
    return table, info


def calc_top(inp: InflationInput) -> dict:
    """Топ-N по м/м, г/г и вкладу в хедлайн на последнюю дату (вклады — по весам Ласпейреса)."""
    latest = inp.latest
    all_cols = inp.meta.index
    mm = inp.values - 100.0
    ew = laspeyres_weights(inp, all_cols)
    headline_weight = ew.loc[latest, inp.cols["headline"]]
    contrib = mm.loc[latest] * ew.loc[latest] / headline_weight
    yoy = yoy_from_mm(mm).loc[latest]
    top: dict = {}
    for group in TOP_GROUPS:
        group_cols = all_cols[inp.meta["group"].eq(str(group))]
        metrics = {"м/м, %": mm.loc[latest, group_cols], "г/г, %": yoy.loc[group_cols],
                   "Вклад в хедлайн, п.п.": contrib.loc[group_cols]}
        top[group] = {}
        for metric, series in metrics.items():
            series = series.copy()
            series.index = inp.meta.loc[series.index, "name"].to_numpy()
            clean = series.replace([np.inf, -np.inf], np.nan).dropna()
            top[group][metric] = (clean[clean > 0].sort_values(ascending=False).head(TOP_N),
                                  clean[clean < 0].sort_values(ascending=True).head(TOP_N))
    return top


def calc_category(inp: InflationInput, name: str) -> dict | None:
    cid = find_category(inp.meta, name)
    if cid is None:
        return None
    mm = (inp.values[cid] - 100.0).loc[:inp.latest]
    table = _cut(pd.DataFrame({"м/м, %": mm, "г/г, %": yoy_from_mm(mm)}), inp.latest)
    if table.empty:
        print(f"⚠️ У категории «{name}» нет данных.")
        return None
    table.insert(0, "Дата", table.index)
    ytd = ytd_by_year(mm)
    if str(inp.latest.year) not in {str(c) for c in ytd.columns}:
        print(f"ℹ️ «{name}»: накопленная с начала {inp.latest.year} г. не считается — "
              f"нет данных за январь или ряд прерывается.")
    ytd.columns = [str(c) for c in ytd.columns]
    ytd.insert(0, "Месяц", [MONTHS_SHORT[m] for m in ytd.index])
    return {"name": str(inp.meta.loc[cid, "name"]), "table": table.reset_index(drop=True),
            "ytd": ytd.reset_index(drop=True), "current_year": str(inp.latest.year)}


def calculate_report(inp: InflationInput) -> ReportData:
    coverage = detailed_coverage(inp)
    hl_mm = inp.values[inp.cols["headline"]] - 100.0
    sa_basis = make_basis(inp, "SA", coverage) if len(inp.sa_values.columns) else None
    if inp.cols.get("headline_sa") is not None:
        hl_sa_mm = inp.values[inp.cols["headline_sa"]] - 100.0
    elif sa_basis is not None:
        hl_sa_mm = weighted_mm(sa_basis.mm, sa_basis.ew)
    else:
        raise ValueError("Нет ни столбца «SA Все товары и услуги», ни SA-данных по категориям.")

    bases = {"SA": sa_basis, "nSA": None}

    def basis_for(kind: str) -> Basis:
        key = "SA" if str(kind).strip().lower() == "sa" else "nSA"
        if bases.get(key) is None:
            bases[key] = make_basis(inp, key, coverage)
        return bases[key]

    core_basis = basis_for(CORE_BASIS)
    core, core_diag = core_measures(core_basis, inp.latest)

    print_checks(inp, coverage, hl_mm, hl_sa_mm, sa_basis, core_basis, core_diag)

    main, main_info = calc_main(inp, hl_mm, hl_sa_mm, core, core_basis)
    contrib, contrib_info = calc_contributions(inp, basis_for(CONTRIBUTIONS_BASIS), hl_mm, hl_sa_mm)
    split, split_info, split_series = calc_split(inp, coverage)

    labels = core_labels()
    derived = {"Хедлайн SA": hl_sa_mm, **split_series}
    derived.update({f"Базовая: {labels[k]}": core[k] for k in CORE_KEYS})
    print("Ряды, доступные для накопленной инфляции (кроме столбцов листа «Данные»): "
          + "; ".join(f"«{k}»" for k in derived))
    cumulative, cumulative_info = calc_cumulative(inp, hl_mm, derived)
    matrix, matrix_info = calc_matrix(inp, basis_for(MATRIX_BASIS), hl_mm, hl_sa_mm)
    top_data = calc_top(inp)
    categories = [c for c in (calc_category(inp, n) for n in SELECTED_CATEGORIES) if c]

    return ReportData(
        latest=inp.latest,
        chart_start=pd.Timestamp(CHART_START),
        main=main, main_info=main_info,
        contrib=contrib, contrib_info=contrib_info,
        cumulative=cumulative, cumulative_info=cumulative_info,
        split=split, split_info=split_info,
        matrix=matrix, matrix_info=matrix_info,
        top_data=top_data, categories=categories,
    )


# =============================================================================
# 6. ПРОВЕРКИ — ПЕЧАТАЮТСЯ В JUPYTER / КОНСОЛЬ
# =============================================================================


def print_checks(inp, coverage, hl_mm, hl_sa_mm, sa_basis, core_basis, core_diag) -> None:
    latest = inp.latest
    meta = inp.meta
    tol = VALIDATION_TOLERANCE_PP
    line = "=" * 78
    print(line)
    print(f"Последнее наблюдение: {MONTHS_NOM[latest.month]} {latest.year} "
          f"(строка весов «Вес {latest.year}»)")

    # 1) Реконструкция официальных индексов из детальных категорий.
    group0 = meta.index[meta["group"].eq("0")]
    if coverage.get(latest, 0) > 0:
        print("-" * 78)
        print("1) Сверка с Росстатом: агрегирование детальных категорий (группа 0) "
              "с весами Ласпейреса")
        detailed = make_basis(inp, "nSA", coverage)
        top_level = {"headline": None, "food": {"ПР"}, "nonfood": {"НЕПР"}, "services": {"У"}}
        titles = {"headline": "Все товары и услуги", "food": "Продовольственные товары",
                  "nonfood": "Непродовольственные товары", "services": "Услуги"}
        start = coverage[coverage > 0].index.min()
        for key, codes in top_level.items():
            cid = inp.cols.get(key)
            if cid is None:
                continue
            cols = group0 if codes is None else group0[_has_any_code(meta.loc[group0, "tokens"], codes).to_numpy()]
            rebuilt = weighted_mm(detailed.mm[cols], detailed.ew[cols])
            error = (rebuilt - (inp.values[cid] - 100.0)).loc[start:latest]
            last = error.get(latest, np.nan)
            status = "✅" if abs(last) <= tol else "⚠️"
            print(f"   {status} {titles[key]:<27} {latest:%m.%Y}: Росстат {inp.values.at[latest, cid] - 100:+.2f}%, "
                  f"расчет {rebuilt.get(latest, np.nan):+.2f}%, расхождение {last:+.3f} п.п.; "
                  f"в среднем с {start:%m.%Y}: {error.abs().mean():.3f} п.п.")
        print(f"   Допуск ±{tol:.2f} п.п. Покрытие весом на последнюю дату: {coverage.get(latest, np.nan):.1f}%.")
        no_weight = meta.loc[group0].index[
            inp.values.loc[str(latest.year), group0].notna().any().to_numpy()
            & inp.weights.loc[latest, group0].isna().to_numpy()
        ]
        if len(no_weight):
            print(f"   ⚠️ Есть данные, но нет веса «Вес {latest.year}» (в расчеты не входят): "
                  f"{meta.loc[no_weight[:8], 'name'].tolist()}{' …' if len(no_weight) > 8 else ''}")
        group1 = meta.index[meta["group"].eq("1")]
        w1 = inp.weights.loc[latest, group1].sum()
        if abs(w1 - 100) > 0.5:
            print(f"   ℹ️ Сумма весов группы «1» в {latest.year} г. = {w1:.2f} (≠ 100): "
                  f"похоже, часть категорий группы 1 входит в другие категории группы 1.")

    # 2) SA-хедлайн.
    if sa_basis is not None:
        print("-" * 78)
        rebuilt_sa = weighted_mm(sa_basis.mm, sa_basis.ew)
        diff = (rebuilt_sa - hl_sa_mm).loc[:latest].abs().max()
        status = "✅" if diff < 0.005 else "⚠️"
        print(f"2) {status} SA-хедлайн, пересчитанный из крупных SA-категорий, vs столбец "
              f"«SA Все товары и услуги»: макс. расхождение {diff:.2e} п.п.")
        weight_sum = inp.weights[sa_basis.cols].where(sa_basis.mm.notna()).sum(axis=1)
        bad_years = sorted({d.year for d, v in weight_sum.loc[:latest].items() if abs(v - 100) > 0.05})
        if bad_years:
            print(f"   ⚠️ Сумма весов крупных SA-категорий ≠ 100 в годах: {bad_years}")
        sa_meta = meta.loc[sa_basis.cols]
        blank = {}
        for key, label in [("code", "Код"), *[(k, v) for k, v in inp.flags.items() if k != "sa"]]:
            empty = sa_meta.index[sa_meta[key].isna()]
            for i in empty:
                blank.setdefault(sa_meta.at[i, "name"], []).append(label)
        for name, fields in blank.items():
            print(f"   ⚠️ «{name}»: не заполнено — {', '.join(fields)}")
        yoy_gap = (yoy_from_mm(hl_sa_mm) - yoy_from_mm(hl_mm)).get(latest, np.nan)
        if abs(yoy_gap) > 0.3:
            print(f"   ⚠️ SA-хедлайн г/г расходится с официальным г/г на {yoy_gap:+.2f} п.п. "
                  f"— за 12 месяцев сезонные факторы должны почти погашаться. Проверьте SA-ряды.")

    # 3) Базовая инфляция.
    print("-" * 78)
    labels = core_labels()
    print(f"3) Базовая инфляция: {core_basis.title}, {latest:%m.%Y}")
    if core_diag.get("codes_not_found"):
        print(f"   ⚠️ Коды исключения не найдены ни у одной категории: {core_diag['codes_not_found']} "
              f"— эти товары/услуги НЕ исключаются.")
    if core_diag.get("n"):
        compact = core_basis.kind == "SA"

        def show(items):
            return ", ".join(items) if compact else f"{len(items)} кат."

        print(f"   • {labels['vol']}: исключено {len(core_diag['vol_dropped'])} из "
              f"{core_diag['vol_eligible']} кат.: {show(core_diag['vol_dropped'])}")
        low = [f"{n} ({v:+.2f})" for n, v in core_diag["trim_low"]]
        high = [f"{n} ({v:+.2f})" for n, v in core_diag["trim_high"]]
        print(f"   • {labels['trim']}: снизу — {show(low)}; сверху — {show(high)}")
        print(f"   • {labels['ex1']}: {core_diag['ex1_share']:.1f}% веса: {show(core_diag['ex1'])}")
        extra = [n for n in core_diag["ex2"] if n not in core_diag["ex1"]]
        print(f"   • {labels['ex2']}: {core_diag['ex2_share']:.1f}% веса; дополнительно: {show(extra)}")
    print(line)


# =============================================================================
# 7. EXCEL: ТАБЛИЦЫ + РЕДАКТИРУЕМЫЕ ГРАФИКИ ПО БРЕНДБУКУ
# =============================================================================


class _Styles:
    def __init__(self, wb):
        text = {"font_name": FONT_TEXT, "font_size": 10, "font_color": BRAND["text"]}
        self.header = wb.add_format({
            "font_name": FONT_TITLE, "font_size": 9, "font_color": "#FFFFFF",
            "bg_color": BRAND["navy"], "text_wrap": True, "align": "center",
            "valign": "vcenter", "border": 1, "border_color": "#FFFFFF",
        })
        self.date = wb.add_format({**text, "num_format": "dd.mm.yyyy", "align": "left"})
        self.num = wb.add_format({**text, "num_format": "0.00"})
        self.text = wb.add_format(text)
        self.title = wb.add_format({"font_name": FONT_TITLE, "font_size": 12,
                                    "font_color": BRAND["text_grey"]})
        self.section = wb.add_format({"font_name": FONT_TITLE, "font_size": 10,
                                      "font_color": "#FFFFFF", "bg_color": BRAND["navy"]})
        self.subheader = wb.add_format({"font_name": FONT_TITLE, "font_size": 9,
                                        "bg_color": BRAND["light_grey"], "align": "center",
                                        "text_wrap": True, "valign": "vcenter"})
        self.note = wb.add_format({**text, "italic": True, "font_size": 9,
                                   "font_color": BRAND["text_grey"]})


def _write_table(ws, st: _Styles, df: pd.DataFrame, row0: int = 0, col0: int = 0,
                 widths: dict | None = None) -> None:
    """Пишет таблицу значениями (без формул). Пустые значения — пустые ячейки."""
    widths = widths or {}
    ws.set_row(row0, 54)
    for j, column in enumerate(df.columns):
        ws.write_string(row0, col0 + j, str(column), st.header)
        if column in ("Дата",):
            width = 11
        elif column in ("Квартал", "Месяц", "Метка"):
            width = 10
        else:
            width = widths.get(column, 13)
        ws.set_column(col0 + j, col0 + j, width)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, value in enumerate(row):
            if _is_na(value):
                continue
            if isinstance(value, (pd.Timestamp, datetime, date)):
                ws.write_datetime(row0 + i, col0 + j, pd.Timestamp(value).to_pydatetime(), st.date)
            elif isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
                ws.write_number(row0 + i, col0 + j, float(value), st.num)
            else:
                ws.write_string(row0 + i, col0 + j, str(value), st.text)


def _font(name=FONT_TEXT, size=9, color=None, **extra) -> dict:
    return {"name": name, "size": size, "bold": False, "color": color or BRAND["text"], **extra}


def _style_chart(chart, title: str, *, y_title: str | None = None, date_axis: bool = True,
                 y_format: str = "0.0", legend: bool = True, size=CHART_SIZE) -> None:
    axis_line = {"color": BRAND["grey"], "width": 0.75}
    chart.set_title({"name": title, "name_font": _font(FONT_TITLE, 11, BRAND["text_grey"]),
                     "overlay": False})
    chart.set_legend({"position": "top", "font": _font()} if legend else {"none": True})
    x_axis = {"num_font": _font(), "line": axis_line, "label_position": "low",
              "major_gridlines": {"visible": False}, "major_tick_mark": "outside"}
    if date_axis:
        x_axis.update({"date_axis": True, "num_format": DATE_AXIS_FORMAT,
                       "major_unit": 1, "major_unit_type": "months",
                       "base_unit": 1, "base_unit_type": "months",
                       "num_font": _font(rotation=-90)})
    else:
        x_axis["major_tick_mark"] = "none"
    chart.set_x_axis(x_axis)
    y_axis = {"num_font": _font(), "num_format": y_format, "line": axis_line,
              "major_tick_mark": "outside",
              "major_gridlines": {"visible": CHART_GRIDLINES,
                                  "line": {"color": BRAND["light_grey"], "width": 0.5}}}
    if y_title:
        y_axis.update({"name": y_title, "name_font": _font(size=9)})
    chart.set_y_axis(y_axis)
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": CHART_BACKGROUND}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"none": True}})
    chart.set_size({"width": size[0], "height": size[1]})


def _line(color, width=1.5, dash="solid") -> dict:
    return {"color": color, "width": width, "dash_type": dash}


def _marker(kind, color, size=5) -> dict:
    return {"type": kind, "size": size, "fill": {"color": color}, "border": {"color": color}}


def _chart_rows(dates: pd.Series, start: pd.Timestamp, row0: int = 0) -> tuple[int, int] | None:
    """Первая и последняя строка листа для графика (данные начинаются со строки row0 + 1)."""
    positions = np.flatnonzero(pd.to_datetime(dates).to_numpy() >= np.datetime64(start))
    if len(positions) == 0:
        return None
    return row0 + 1 + int(positions[0]), row0 + len(dates)


def _sheet_main(wb, st, report: ReportData) -> None:
    name = "Хедлайн и базовая"
    ws = wb.add_worksheet(name)
    df, info = report.main, report.main_info
    _write_table(ws, st, df)
    ws.freeze_panes(1, 1)
    rows = _chart_rows(df["Дата"], report.chart_start)
    if rows is None:
        return
    r1, r2 = rows
    col = {c: j for j, c in enumerate(df.columns)}
    labels = {k: v[0].upper() + v[1:] for k, v in info["labels"].items()}
    cats = [name, r1, 0, r2, 0]

    def ref(column):
        return [name, r1, col[column], r2, col[column]]

    # (1) м/м: SA-хедлайн и базовая — ярко, nSA-хедлайн — бледно.
    ch = wb.add_chart({"type": "line"})
    ch.add_series({"name": "Хедлайн nSA", "categories": cats, "values": ref("Хедлайн nSA, м/м, %"),
                   "line": _line(HEADLINE_NSA_COLOR, 1.5)})
    for k in CORE_KEYS:
        color, dash, _ = CORE_STYLES[k]
        ch.add_series({"name": labels[k], "categories": cats,
                       "values": ref(f"Базовая {info['kind']}: {info['labels'][k]}, м/м, %"),
                       "line": _line(color, 1.75, dash)})
    ch.add_series({"name": "Хедлайн SA", "categories": cats, "values": ref("Хедлайн SA, м/м, %"),
                   "line": _line(HEADLINE_COLOR, 2.25)})
    core_note = "SA" if info["kind"] == "SA" else "nSA, детальные категории"
    _style_chart(ch, f"Инфляция м/м: хедлайн и базовая ({core_note}), %")
    first_chart_col = len(df.columns) + 1
    ws.insert_chart(1, first_chart_col, ch)

    # (2) г/г — линии, SAAR — точки.
    ch = wb.add_chart({"type": "line"})
    for k in CORE_KEYS:
        color, dash, _ = CORE_STYLES[k]
        ch.add_series({"name": f"{labels[k]}, г/г", "categories": cats,
                       "values": ref(f"Базовая: {info['labels'][k]}, г/г, %"),
                       "line": _line(color, 1.75, dash)})
    ch.add_series({"name": "Хедлайн, г/г", "categories": cats, "values": ref("Хедлайн, г/г, %"),
                   "line": _line(HEADLINE_COLOR, 2.25)})
    ch.add_series({"name": "Хедлайн, SAAR", "categories": cats, "values": ref("Хедлайн SAAR, %"),
                   "line": {"none": True}, "marker": _marker("circle", HEADLINE_COLOR, 6)})
    if info["kind"] == "SA":
        for k in CORE_KEYS:
            color, _, marker = CORE_STYLES[k]
            ch.add_series({"name": f"{labels[k]}, SAAR", "categories": cats,
                           "values": ref(f"Базовая: {info['labels'][k]}, SAAR, %"),
                           "line": {"none": True}, "marker": _marker(marker, color, 5)})
    title = "Инфляция г/г (линии) и SAAR (точки): хедлайн и базовая, %"
    if info["kind"] != "SA":
        title = "Инфляция г/г (линии) и SAAR хедлайна (точки), %"
    _style_chart(ch, title, size=(CHART_SIZE[0], CHART_SIZE[1] + 40))
    ws.insert_chart(24, first_chart_col, ch)


def _sheet_contributions(wb, st, report: ReportData) -> None:
    df, info = report.contrib, report.contrib_info
    name = "Вклады"
    ws = wb.add_worksheet(name)
    _write_table(ws, st, df, widths={c: 14 for c in df.columns})
    ws.freeze_panes(1, 3)
    if df.empty:
        return
    # График начинается с первого месяца квартала, в который попадает CHART_START.
    start = report.chart_start
    start = pd.Timestamp(start.year, 3 * ((start.month - 1) // 3) + 1, 1)
    rows = _chart_rows(df["Дата"].fillna(pd.Timestamp("1900-01-01")), start)
    if rows is None:
        return
    r1, r2 = rows
    cat_ref = [name, r1, 0, r2, 1]
    part = df.iloc[r1 - 1:r2]
    # Двухуровневая ось: месяцы и под ними кварталы. Кэш подписей — в формате Excel:
    # пустые ячейки пропускаются (подпись квартала «растягивается» до следующей).
    def labels(column):
        return [v if isinstance(v, str) and v != "" else None for v in part[column]]

    cats_cache = [labels("Квартал"), labels("Месяц")]
    col = {c: j for j, c in enumerate(df.columns)}

    ch = wb.add_chart({"type": "column", "subtype": "stacked"})
    for i, component in enumerate(info["components"]):
        color = PALETTE_SIMPLE[i % len(PALETTE_SIMPLE)]
        if component.startswith(("Вклад: Не распределено", "Расхождение")):
            color = BRAND["light_grey"]
        label = component.replace("Вклад: ", "").replace(", п.п.", "")
        series = {"name": label, "categories": cat_ref, "categories_data": cats_cache,
                  "values": [name, r1, col[component], r2, col[component]],
                  "fill": {"color": color}, "border": {"none": True}}
        if i == 0:
            series["gap"] = 40
        ch.add_series(series)
    line = wb.add_chart({"type": "line"})
    line.add_series({"name": info["line"].replace(", м/м, %", " м/м"),
                     "categories": cat_ref, "categories_data": cats_cache,
                     "values": [name, r1, col[info["line"]], r2, col[info["line"]]],
                     "line": _line(BRAND["text"], 1.25),
                     "marker": _marker("circle", BRAND["text"], 4)})
    ch.combine(line)
    ch.show_blanks_as("span")
    title = ("Вклады в SA-хедлайн м/м, п.п." if info["basis"].kind == "SA"
             else "Вклады в хедлайн м/м (nSA, детальные категории), п.п.")
    _style_chart(ch, title, date_axis=False)
    ws.insert_chart(1, len(df.columns) + 1, ch)


def _sheet_cumulative(wb, st, report: ReportData) -> None:
    df, info = report.cumulative, report.cumulative_info
    name = "Накопленная"
    ws = wb.add_worksheet(name)
    _write_table(ws, st, df, widths={c: 16 for c in df.columns})
    ws.freeze_panes(1, 1)
    rows = _chart_rows(df["Дата"], report.chart_start)
    if rows is None or len(df.columns) < 2:
        return
    r1, r2 = rows
    ch = wb.add_chart({"type": "line"})
    for i, column in enumerate(df.columns[1:]):
        color = PALETTE_SIMPLE[i % len(PALETTE_SIMPLE)]
        ch.add_series({"name": column.split(", % к ")[0], "categories": [name, r1, 0, r2, 0],
                       "values": [name, r1, i + 1, r2, i + 1],
                       "line": _line(color, 2.25 if i == 0 else 1.75)})
    _style_chart(ch, f"Накопленная инфляция: уровень цен, % к {info['base_label']}",
                 legend=len(df.columns) > 2)
    ws.insert_chart(1, len(df.columns) + 1, ch)


def _sheet_split(wb, st, report: ReportData) -> None:
    df, info = report.split, report.split_info
    n1, n0 = info["names"]
    name = f"{n1} и {n0.lower()}"[:31]
    ws = wb.add_worksheet(name)
    _write_table(ws, st, df, widths={c: 15 for c in df.columns})
    ws.freeze_panes(1, 1)
    measure = "м/м" if normalize_text(SPLIT_CHART) in ("м/м", "мм") else "г/г"
    # Г/г появляется только через 12 месяцев после начала детальных данных:
    # график начинается с первого месяца, где есть значения.
    shown = df[[f"{n1}, {measure}, %", f"{n0}, {measure}, %"]].notna().any(axis=1)
    first_valid = df.loc[shown, "Дата"].min()
    start = report.chart_start if pd.isna(first_valid) else max(report.chart_start, first_valid)
    rows = _chart_rows(df["Дата"], start)
    if rows is None:
        return
    r1, r2 = rows
    ch = wb.add_chart({"type": "line"})
    for series_name, color in [(n1, BRAND["navy"]), (n0, BRAND["magenta"])]:
        j = list(df.columns).index(f"{series_name}, {measure}, %")
        ch.add_series({"name": series_name, "categories": [name, r1, 0, r2, 0],
                       "values": [name, r1, j, r2, j], "line": _line(color, 2.25)})
    _style_chart(ch, f"{n1} и {n0.lower()} инфляция, {measure}, %")
    ws.insert_chart(1, len(df.columns) + 1, ch)


def _sheet_matrix(wb, st, report: ReportData) -> None:
    df, info = report.matrix, report.matrix_info
    name = "Матрица"
    ws = wb.add_worksheet(name)
    _write_table(ws, st, df, widths={c: 15 for c in df.columns})
    ws.freeze_panes(1, 1)
    rows = _chart_rows(df["Дата"], info["window_start"])
    if rows is None:
        return
    r1, r2 = rows
    col = {c: j for j, c in enumerate(df.columns)}
    part = df.iloc[r1 - 1:r2]
    tier_style = [(BRAND["navy"], 9, BRAND["navy"]),
                  (BRAND["grey_blue"], 7, BRAND["grey_blue"]),
                  (BRAND["light_grey"], 6, None)]
    ch = wb.add_chart({"type": "scatter"})
    for t, tier_name in enumerate(info["tiers"]):
        color, size, label_color = tier_style[t]
        if part[tier_name].notna().sum() == 0:
            continue  # например, в январе нет точек «ранее в текущем году»
        series = {"name": tier_name,
                  "categories": [name, r1, col[info["x"]], r2, col[info["x"]]],
                  "values": [name, r1, col[tier_name], r2, col[tier_name]],
                  "marker": _marker("circle", color, size)}
        if label_color is not None:
            custom = [{"value": lab, "font": _font(size=9, color=label_color)}
                      if not _is_na(v) else {"delete": True}
                      for lab, v in zip(part["Метка"], part[tier_name])]
            series["data_labels"] = {"value": True, "custom": custom, "position": "right"}
        ch.add_series(series)
    axis_line = {"color": BRAND["grey"], "width": 0.75}
    ch.set_title({"name": "Инфляция и разброс изменений цен по корзине",
                  "name_font": _font(FONT_TITLE, 11, BRAND["text_grey"]), "overlay": False})
    ch.set_legend({"position": "top", "font": _font()})
    # Оси пересекаются на границах «высокая/низкая» и «однородная/неоднородная»:
    # получается решетка из четырех квадрантов, подписи осей — по краям.
    ch.set_x_axis({"name": info["x"], "name_font": _font(), "num_font": _font(),
                   "num_format": "0.0", "line": axis_line, "label_position": "low",
                   "crossing": round(info["x_threshold"], 3),
                   "major_gridlines": {"visible": False}})
    ch.set_y_axis({"name": info["y"], "name_font": _font(), "num_font": _font(),
                   "num_format": "0.0", "line": axis_line, "label_position": "low",
                   "crossing": round(info["y_threshold"], 3),
                   "major_gridlines": {"visible": False}})
    ch.set_chartarea({"border": {"none": True}, "fill": {"color": CHART_BACKGROUND}})
    ch.set_plotarea({"border": {"none": True}, "fill": {"none": True}})
    ch.set_size({"width": 720, "height": 560})
    ws.insert_chart(1, len(df.columns) + 1, ch)
    note_row = 30
    ws.write(note_row, len(df.columns) + 1,
             f"Горизонтальная ось пересекает вертикальную на уровне {info['y_threshold']:.2f}% м/м "
             f"(≈ {MATRIX_TARGET_SAAR:.0f}% в год), вертикальная — на медиане разброса "
             f"за период графика ({info['x_threshold']:.2f} п.п.).", st.note)


def _sheet_top(wb, st, report: ReportData) -> None:
    ws = wb.add_worksheet("Топ-5")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 7, st.text)
    ws.set_column("B:B", 46, st.text)
    ws.set_column("C:C", 12, st.num)
    ws.set_column("D:D", 3)
    ws.set_column("E:E", 7, st.text)
    ws.set_column("F:F", 46, st.text)
    ws.set_column("G:G", 12, st.num)
    ws.write("A1", f"Топ-{TOP_N} категорий, {MONTHS_NOM[report.latest.month]} {report.latest.year}", st.title)
    row = 2
    for group in TOP_GROUPS:
        ws.merge_range(row, 0, row, 6, f"Группа {group}", st.section)
        row += 1
        for metric in ["м/м, %", "г/г, %", "Вклад в хедлайн, п.п."]:
            positive, negative = report.top_data[str(group)][metric]
            ws.merge_range(row, 0, row, 6, metric, st.subheader)
            row += 1
            ws.write_row(row, 0, ["Место", "Сильнее всего выросли", "Значение"], st.subheader)
            ws.write_row(row, 4, ["Место", "Сильнее всего снизились", "Значение"], st.subheader)
            row += 1
            for rank in range(TOP_N):
                if rank < len(positive):
                    ws.write_number(row + rank, 0, rank + 1, st.text)
                    ws.write_string(row + rank, 1, str(positive.index[rank]), st.text)
                    ws.write_number(row + rank, 2, float(positive.iloc[rank]), st.num)
                if rank < len(negative):
                    ws.write_number(row + rank, 4, rank + 1, st.text)
                    ws.write_string(row + rank, 5, str(negative.index[rank]), st.text)
                    ws.write_number(row + rank, 6, float(negative.iloc[rank]), st.num)
            row += TOP_N + 2
    ws.write(row, 0, "Вклад — по весам Ласпейреса (годовой вес × накопленный с декабря индекс).", st.note)


def _safe_sheet_name(name: str, used: set) -> str:
    clean = re.sub(r"[\[\]:*?/\\]", " ", name).strip()[:31] or "Категория"
    candidate, k = clean, 2
    while candidate.lower() in used:
        suffix = f" ({k})"
        candidate = clean[:31 - len(suffix)] + suffix
        k += 1
    used.add(candidate.lower())
    return candidate


def _sheet_category(wb, st, report: ReportData, cat: dict, used: set) -> None:
    name = _safe_sheet_name(cat["name"], used)
    ws = wb.add_worksheet(name)
    table, ytd = cat["table"], cat["ytd"]
    _write_table(ws, st, table)
    ws.freeze_panes(1, 1)
    ytd_col = len(table.columns) + 1
    _write_table(ws, st, ytd, col0=ytd_col, widths={c: 8 for c in ytd.columns})
    chart_col = ytd_col + len(ytd.columns) + 1

    rows = _chart_rows(table["Дата"], report.chart_start)
    if rows is not None:
        r1, r2 = rows
        ch = wb.add_chart({"type": "line"})
        ch.add_series({"name": "г/г", "categories": [name, r1, 0, r2, 0],
                       "values": [name, r1, 2, r2, 2], "line": _line(BRAND["navy"], 2.25)})
        ch.add_series({"name": "м/м (правая шкала)", "categories": [name, r1, 0, r2, 0],
                       "values": [name, r1, 1, r2, 1], "line": _line(BRAND["magenta"], 1.5),
                       "y2_axis": True})
        _style_chart(ch, f"{cat['name']}: м/м и г/г, %")
        ch.set_y2_axis({"num_font": _font(), "num_format": "0.0",
                        "line": {"color": BRAND["grey"], "width": 0.75},
                        "major_gridlines": {"visible": False}})
        ws.insert_chart(1, chart_col, ch)

    # Накопленная с начала года: текущий год — ярко, прошлый — темно-синим,
    # более ранние годы — светло-серым (в легенде только текущий и прошлый).
    years = list(ytd.columns[1:])
    if years:
        ch = wb.add_chart({"type": "line"})
        order = [y for y in years if y != cat["current_year"]] + (
            [cat["current_year"]] if cat["current_year"] in years else [])
        previous = str(int(cat["current_year"]) - 1)
        hidden = []
        for i, year in enumerate(order):
            j = ytd_col + list(ytd.columns).index(year)
            if year == cat["current_year"]:
                style = {"line": _line(BRAND["magenta"], 2.75), "marker": _marker("circle", BRAND["magenta"], 5)}
            elif year == previous:
                style = {"line": _line(BRAND["navy"], 1.75)}
            else:
                style = {"line": _line(BRAND["light_grey"], 1.0)}
                hidden.append(i)
            ch.add_series({"name": year, "categories": [name, 1, ytd_col, 12, ytd_col],
                           "values": [name, 1, j, 12, j], **style})
        _style_chart(ch, f"{cat['name']}: накопленная с начала года инфляция, %",
                     date_axis=False)
        if hidden:
            ch.set_legend({"position": "top", "font": _font(), "delete_series": hidden})
        ws.insert_chart(24, chart_col, ch)


def export_excel(report: ReportData, path: Path) -> None:
    import xlsxwriter
    from xlsxwriter.exceptions import FileCreateError

    path = Path(path)
    wb = xlsxwriter.Workbook(str(path))
    wb.formats[0].set_font_name(FONT_TEXT)
    st = _Styles(wb)
    _sheet_main(wb, st, report)
    _sheet_contributions(wb, st, report)
    _sheet_cumulative(wb, st, report)
    _sheet_split(wb, st, report)
    _sheet_matrix(wb, st, report)
    _sheet_top(wb, st, report)
    used = {ws.get_name().lower() for ws in wb.worksheets()}
    for cat in report.categories:
        _sheet_category(wb, st, report, cat, used)
    try:
        wb.close()
    except (PermissionError, FileCreateError) as exc:
        raise PermissionError(f"Не удалось записать {path}: если файл открыт в Excel, "
                              f"закройте его и запустите снова.") from exc
    print(f"✅ Excel сохранен: {path.resolve()}")


# =============================================================================
# 8. PDF: ОПЦИОНАЛЬНО, ПО ОДНОМУ ГРАФИКУ НА СТРАНИЦУ
# =============================================================================


def export_pdf(report: ReportData, path: Path) -> None:
    import matplotlib
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    from matplotlib.ticker import FuncFormatter

    matplotlib.rcParams.update({
        "font.family": "serif",
        "font.serif": [FONT_TEXT, "Liberation Serif", "DejaVu Serif"],
        "axes.spines.top": False, "axes.spines.right": False,
        "axes.edgecolor": BRAND["grey"], "axes.linewidth": 0.75,
        "xtick.color": BRAND["text"], "ytick.color": BRAND["text"],
        "legend.frameon": False, "axes.formatter.use_locale": False,
    })
    title_font = {"family": "sans-serif", "fontsize": 13, "color": BRAND["text_grey"],
                  "fontweight": "normal"}
    comma = FuncFormatter(lambda v, _: f"{v:.1f}".replace(".", ","))

    def date_axis(ax, dates):
        dates = pd.to_datetime(pd.Series(dates)).dropna()
        ticks = [d for d in dates if d.month % 2 == 1] or list(dates)
        ax.set_xticks(ticks)
        ax.set_xticklabels([f"{MONTHS_SHORT[d.month]} {d.year % 100:02d}" for d in ticks], rotation=90)

    def finish(fig, ax, title, legend=True):
        ax.set_title(title, loc="left", **title_font)
        ax.yaxis.set_major_formatter(comma)
        if legend:
            ax.legend(loc="upper left", bbox_to_anchor=(0, 1.0), ncol=3, fontsize=9)
        fig.tight_layout()
        pdf.savefig(fig)
        plt.close(fig)

    start = report.chart_start
    with PdfPages(path) as pdf:
        # Лист 1.
        info = report.main_info
        labels = {k: v[0].upper() + v[1:] for k, v in info["labels"].items()}
        d = report.main[report.main["Дата"] >= start]
        fig, ax = plt.subplots(figsize=(11.7, 6.8))
        ax.plot(d["Дата"], d["Хедлайн nSA, м/м, %"], color=HEADLINE_NSA_COLOR, lw=1.5, label="Хедлайн nSA")
        for k in CORE_KEYS:
            color, dash, _ = CORE_STYLES[k]
            ax.plot(d["Дата"], d[f"Базовая {info['kind']}: {info['labels'][k]}, м/м, %"], color=color,
                    lw=1.75, ls="--" if dash == "dash" else "-", label=labels[k])
        ax.plot(d["Дата"], d["Хедлайн SA, м/м, %"], color=HEADLINE_COLOR, lw=2.25, label="Хедлайн SA")
        date_axis(ax, d["Дата"])
        finish(fig, ax, "Инфляция м/м: хедлайн и базовая, %")

        fig, ax = plt.subplots(figsize=(11.7, 6.8))
        markers = {"circle": "o", "square": "s", "diamond": "D", "triangle": "^"}
        for k in CORE_KEYS:
            color, dash, marker = CORE_STYLES[k]
            ax.plot(d["Дата"], d[f"Базовая: {info['labels'][k]}, г/г, %"], color=color, lw=1.75,
                    ls="--" if dash == "dash" else "-", label=f"{labels[k]}, г/г")
            column = f"Базовая: {info['labels'][k]}, SAAR, %"
            if column in d:
                ax.scatter(d["Дата"], d[column], color=color, s=18, marker=markers[marker],
                           label=f"{labels[k]}, SAAR", zorder=3)
        ax.plot(d["Дата"], d["Хедлайн, г/г, %"], color=HEADLINE_COLOR, lw=2.25, label="Хедлайн, г/г")
        ax.scatter(d["Дата"], d["Хедлайн SAAR, %"], color=HEADLINE_COLOR, s=24, label="Хедлайн, SAAR", zorder=3)
        date_axis(ax, d["Дата"])
        finish(fig, ax, "Инфляция г/г (линии) и SAAR (точки), %")

        # Лист 2: вклады, месяцы по кварталам.
        c, cinfo = report.contrib, report.contrib_info
        q_start = pd.Timestamp(start.year, 3 * ((start.month - 1) // 3) + 1, 1)
        first = c.index[c["Дата"].ge(q_start).fillna(False)]
        if len(first):
            part = c.loc[first[0]:].reset_index(drop=True)
            x = np.arange(len(part))
            fig, ax = plt.subplots(figsize=(11.7, 6.8))
            pos, neg = np.zeros(len(part)), np.zeros(len(part))
            for i, component in enumerate(cinfo["components"]):
                vals = part[component].astype(float).fillna(0).to_numpy()
                color = PALETTE_SIMPLE[i % len(PALETTE_SIMPLE)]
                if component.startswith(("Вклад: Не распределено", "Расхождение")):
                    color = BRAND["light_grey"]
                bottom = np.where(vals >= 0, pos, neg)
                ax.bar(x, vals, bottom=bottom, color=color, width=0.8,
                       label=component.replace("Вклад: ", "").replace(", п.п.", ""))
                pos += np.where(vals > 0, vals, 0)
                neg += np.where(vals < 0, vals, 0)
            total = part[cinfo["line"]].astype(float)
            ok = total.notna().to_numpy()
            ax.plot(x[ok], total[ok], color=BRAND["text"], lw=1.25, marker="o", ms=3,
                    label=cinfo["line"].replace(", м/м, %", " м/м"))
            ax.axhline(0, color=BRAND["grey"], lw=0.75)
            ax.set_xticks(x)
            ax.set_xticklabels(part["Месяц"].fillna(""), fontsize=8)
            for i, q in enumerate(part["Квартал"].fillna("")):
                if q.strip():
                    n_months = int(part["Месяц"].iloc[i:i + 3].fillna("").astype(bool).sum())
                    ax.annotate(q, xy=(i + (n_months - 1) / 2, 0), xycoords=("data", "axes fraction"),
                                xytext=(0, -24), textcoords="offset points", ha="center", fontsize=9)
            finish(fig, ax, "Вклады в хедлайн м/м, п.п.")

        # Лист 3: накопленная.
        cu = report.cumulative
        d = cu[cu["Дата"] >= start]
        fig, ax = plt.subplots(figsize=(11.7, 6.8))
        for i, column in enumerate(cu.columns[1:]):
            ax.plot(d["Дата"], d[column], color=PALETTE_SIMPLE[i % len(PALETTE_SIMPLE)],
                    lw=2.25 if i == 0 else 1.75, label=column.split(", % к ")[0])
        ax.axhline(0, color=BRAND["grey"], lw=0.75)
        date_axis(ax, d["Дата"])
        finish(fig, ax, f"Накопленная инфляция: уровень цен, % к {report.cumulative_info['base_label']}",
               legend=len(cu.columns) > 2)

        # Лист 4: монетарная / немонетарная.
        sp = report.split
        n1, n0 = report.split_info["names"]
        measure = "м/м" if normalize_text(SPLIT_CHART) in ("м/м", "мм") else "г/г"
        shown = sp[[f"{n1}, {measure}, %", f"{n0}, {measure}, %"]].notna().any(axis=1)
        split_start = max(start, sp.loc[shown, "Дата"].min()) if shown.any() else start
        d = sp[sp["Дата"] >= split_start]
        fig, ax = plt.subplots(figsize=(11.7, 6.8))
        ax.plot(d["Дата"], d[f"{n1}, {measure}, %"], color=BRAND["navy"], lw=2.25, label=n1)
        ax.plot(d["Дата"], d[f"{n0}, {measure}, %"], color=BRAND["magenta"], lw=2.25, label=n0)
        date_axis(ax, d["Дата"])
        finish(fig, ax, f"{n1} и {n0.lower()} инфляция, {measure}, %")

        # Лист 5: матрица.
        m, minfo = report.matrix, report.matrix_info
        d = m[m["Дата"] >= minfo["window_start"]]
        fig, ax = plt.subplots(figsize=(9.5, 7.5))
        styles = [(BRAND["navy"], 70, True), (BRAND["grey_blue"], 45, True), (BRAND["light_grey"], 30, False)]
        for t, tier in reversed(list(enumerate(minfo["tiers"]))):
            color, size, label_points = styles[t]
            sel = d[tier].notna()
            ax.scatter(d.loc[sel, minfo["x"]], d.loc[sel, tier], color=color, s=size, label=tier, zorder=3 - t)
            if label_points:
                for _, r in d[sel].iterrows():
                    ax.annotate(r["Метка"], (r[minfo["x"]], r[tier]), xytext=(5, 0),
                                textcoords="offset points", fontsize=9, color=color, va="center")
        ax.axhline(minfo["y_threshold"], color=BRAND["grey"], lw=0.75)
        ax.axvline(minfo["x_threshold"], color=BRAND["grey"], lw=0.75)
        ax.set_xlabel(minfo["x"])
        ax.set_ylabel(minfo["y"])
        ax.xaxis.set_major_formatter(comma)
        finish(fig, ax, "Инфляция и разброс изменений цен по корзине")

        # Листы 7+: категории.
        for cat in report.categories:
            t = cat["table"]
            d = t[t["Дата"] >= start]
            fig, ax = plt.subplots(figsize=(11.7, 6.8))
            ax.plot(d["Дата"], d["г/г, %"], color=BRAND["navy"], lw=2.25, label="г/г")
            ax2 = ax.twinx()
            ax2.plot(d["Дата"], d["м/м, %"], color=BRAND["magenta"], lw=1.5, label="м/м (правая шкала)")
            ax2.spines["right"].set_visible(True)
            ax2.yaxis.set_major_formatter(comma)
            lines = ax.get_lines() + ax2.get_lines()
            ax.legend(lines, [ln.get_label() for ln in lines], loc="upper left", fontsize=9)
            date_axis(ax, d["Дата"])
            finish(fig, ax, f"{cat['name']}: м/м и г/г, %", legend=False)

            ytd = cat["ytd"]
            if len(ytd.columns) < 2:
                continue
            fig, ax = plt.subplots(figsize=(11.7, 6.8))
            for year in ytd.columns[1:]:
                if year == cat["current_year"]:
                    continue
                prev = year == str(int(cat["current_year"]) - 1)
                ax.plot(range(12), ytd[year], color=BRAND["navy"] if prev else BRAND["light_grey"],
                        lw=1.75 if prev else 1.0, label=year if prev else None)
            if cat["current_year"] in ytd:
                ax.plot(range(12), ytd[cat["current_year"]], color=BRAND["magenta"], lw=2.75,
                        marker="o", ms=4, label=cat["current_year"])
            ax.set_xticks(range(12))
            ax.set_xticklabels(ytd["Месяц"])
            finish(fig, ax, f"{cat['name']}: накопленная с начала года инфляция, %")

    print(f"✅ PDF сохранен: {Path(path).resolve()}")


# =============================================================================
# 9. ЗАПУСК
# =============================================================================


def main() -> ReportData:
    inp = load_input(INPUT_FILE)
    report = calculate_report(inp)
    export_excel(report, OUTPUT_XLSX)
    if EXPORT_PDF:
        export_pdf(report, OUTPUT_PDF)
    return report


# В Jupyter: %run inflation_report.py — результат будет в переменной report.
if __name__ == "__main__":
    report = main()
