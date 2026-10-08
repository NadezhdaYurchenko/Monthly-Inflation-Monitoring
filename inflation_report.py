# %%
"""
Месячный мониторинг инфляции.

Скрипт читает исходный Excel-файл, считает показатели в Python и сохраняет отчет
в Excel: на каждом листе таблица с данными (за весь период, с 2002 г.) и рядом
редактируемые Excel-графики (с 2024 г.). По желанию дополнительно формирует PDF
с теми же графиками.

Листы отчета:
    1. Хедлайн и базовая — м./м. и г./г. хедлайна, диапазон четырех мер базовой
       инфляции, SAAR;
    2. Вклады — разложение SA-хедлайна м./м. на вклады (набор компонентов задается
       в настройках), месяцы на графике сгруппированы по кварталам;
    3. Накопленная — накопленная с начала года инфляция: текущий год против
       прошлых лет;
    4. Монетарная и немонетарная инфляция, м./м. и г./г.;
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
#   "SA"  — крупные категории с листа «SA крупные категории». И отбор категорий
#           (волатильность, усечение), и агрегирование — по SA-данным.
#           История с 2002 г. Рекомендуемый вариант.
#   "nSA" — детальные категории (группа 0) без сезонной корректировки.
#           История только с 2024 г.; SAAR базовой инфляции не считается.
CORE_BASIS = "SA"

# Метод 1 — без 20% наиболее волатильных категорий.
CORE_VOLATILITY_WINDOW = 3      # окно в месяцах (включая текущий) для дисперсии м./м.
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

# ── 4. Лист 2: вклады в хедлайн м./м. ─────────────────────────────────────────
# "SA"  — вклады в SA-хедлайн по крупным SA-категориям (история с 2002 г.).
#         Сумма вкладов точно равна столбцу «SA Все товары и услуги».
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
# «Прочие (для SA)» всегда относятся к своей группе (ПР / НЕПР / У).
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
# CONTRIBUTIONS = [                        # у «Прочих (для SA)» флагов нет — они в «остальном»
#     ("Подверженные колебаниям валютного курса", {"Вал.курс": 1}),
#     ("Остальные товары и услуги", {"остальное": True}),
# ]
# CONTRIBUTIONS = [                        # только с CONTRIBUTIONS_BASIS = "nSA"
#     ("Продовольственные товары", {"код": ["ПР"]}),
#     ("Непродовольственные товары", {"код": ["НЕПР"]}),
#     ("Регулируемые услуги", {"код": ["РУ"]}),
#     ("Нерегулируемые услуги", {"код": ["У"]}),
# ]

# ── 5. Лист 3: накопленная с начала года инфляция ─────────────────────────────
# Хедлайн по годам: текущий год — розовым, прошлый — темно-синим и еще столько
# лет до прошлого — бледными цветами. В таблице — все годы с 2002 г.
CUMULATIVE_PALE_YEARS = 5

# ── 6. Лист 4: монетарная и немонетарная инфляция ─────────────────────────────
# Детальные категории делятся по флагу: 1 → первая линия, 0 → вторая.
# Два графика: м./м. и г./г.
SPLIT_FLAG = "Монетар. Инфляция"
SPLIT_NAMES = ("Монетарная", "Немонетарная")

# ── 7. Лист 5: матрица «инфляция — разброс по корзине» ────────────────────────
MATRIX_BASIS = "SA"             # "SA":  SA-хедлайн м./м. и разброс SA м./м. по крупным категориям
                                # "nSA": хедлайн м./м. и разброс м./м. по детальным категориям
MATRIX_DISPERSION = "std"       # "std" — взвешенное стандартное отклонение м./м. по корзине;
                                # "iqr" — взвешенный межквартильный размах (устойчив к выбросам;
                                #         для "nSA" лучше он: иначе разброс задают единичные позиции)
MATRIX_YEARS = 3                # сколько календарных лет наносить (включая текущий)
MATRIX_TARGET_SAAR = 4.0        # граница «высокая/низкая»: м./м., соответствующий 4% в год

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
VALIDATION_TOLERANCE_PP = 0.15          # допуск сверки с Росстатом, п.п. м./м.
MIN_DETAILED_COVERAGE_WEIGHT = 95.0     # месяцы с меньшим покрытием веса детальными
                                        # категориями не считаются (в % веса)
COLUMN_ALIASES = {
    "headline": ["Все товары и услуги"],
    "headline_sa": ["SA Все товары и услуги"],
    "food": ["Продовольственные товары"],
    "nonfood": ["Непродовольственные товары"],
    "services": ["Услуги"],
    "sa_food": ["SA Продовольственые товары", "SA Продовольственные товары"],
    "sa_nonfood": ["SA Непродовольственная инфляция", "SA Непродовольственные товары"],
    "sa_services": ["SA Услуги"],
}

# ╔═════════════════════════════════════════════════════════════════════════════╗
# ║                               ОФОРМЛЕНИЕ                                    ║
# ╚═════════════════════════════════════════════════════════════════════════════╝
# Цвета — брендбук, раздел 1.9 «Цвета для графиков и диаграмм». Остальное — как
# принято в центре: всё набирается Arial; заголовок графика слева, черный,
# жирный; оси черные; ось Y справа (пересекает ось X «в максимальном значении»);
# легенда снизу; без линий сетки; на линиях хедлайна нет маркеров; боковой зазор
# у столбцов 20%.
BRAND = {
    "navy": "#1E3B56", "magenta": "#B1046E", "beige": "#A99892", "blue": "#009AD9",
    "coral": "#ED695A", "purple": "#6758A2", "green": "#6EBC84", "mustard": "#D5AD00",
    "grey_blue": "#818DA2", "pink": "#D492B6", "light_beige": "#D4CAC7",
    "light_blue": "#A2CCEE", "light_coral": "#F7BAAA", "light_purple": "#B1A8D3",
    "light_green": "#C0DFC3", "light_mustard": "#EBD592",
    "grey": "#A8B6BF", "light_grey": "#D1DADF",
}
# Порядок добавления цветов на графиках (брендбук, 1.9).
PALETTE_SIMPLE = [BRAND[k] for k in (
    "navy", "magenta", "beige", "blue", "grey_blue", "pink", "light_beige", "light_blue",
    "coral", "purple", "green", "mustard", "light_coral", "light_purple", "light_green",
    "light_mustard",
)]
FONT = "Arial"
FONT_SIZE = 10                          # подписи осей, легенда, таблицы
TITLE_SIZE = 12                         # заголовки графиков
AXIS_COLOR = "#000000"
CHART_BACKGROUND = "#FFFFFF"
CHART_GRIDLINES = False
CHART_SIZE = (950, 480)                 # ширина и высота графиков в пикселях
COLUMN_GAP = 20                         # боковой зазор между столбцами, %
DATE_AXIS_FORMAT = "[$-419]mmm yy"      # «янв. 24» при любом языке Excel

HEADLINE_COLOR = BRAND["navy"]
HEADLINE_NSA_COLOR = BRAND["grey"]
CORE_BAND_COLOR = BRAND["pink"]         # диапазон базовой инфляции
CORE_BAND_TRANSPARENCY = 40             # прозрачность заливки диапазона, %
# Точки SAAR базовой инфляции: цвет и форма маркера.
CORE_SAAR_MARKERS = {
    "vol": (BRAND["magenta"], "circle"),
    "trim": (BRAND["pink"], "square"),
    "ex1": (BRAND["magenta"], "diamond"),
    "ex2": (BRAND["pink"], "triangle"),
}
CURRENT_YEAR_COLOR = BRAND["magenta"]
PREVIOUS_YEAR_COLOR = BRAND["navy"]
# Бледные годы на графике накопленной инфляции: одна голубо-сиреневая шкала между
# фирменными #B1A8D3 и #A2CCEE — чем старше год, тем светлее (от ближнего к дальнему).
PALE_YEAR_COLORS = ["#948CBD", "#A3A0CE", "#A6B7E0", "#ACCDEB", "#C2E0F2"]
# Матрица: последние 3 месяца, остальной текущий год, прошлые годы.
MATRIX_COLORS = (BRAND["magenta"], BRAND["navy"], BRAND["light_grey"])
MATRIX_LABEL_COLORS = (BRAND["magenta"], BRAND["navy"], BRAND["grey"])

# =============================================================================
#                        ДАЛЕЕ — КОД (менять не нужно)
# =============================================================================

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Iterable
import difflib
import math
import re
import textwrap

import numpy as np
import pandas as pd


MM, YY = "м./м.", "г./г."
MONTHS_SHORT = {1: "янв", 2: "фев", 3: "мар", 4: "апр", 5: "май", 6: "июн",
                7: "июл", 8: "авг", 9: "сен", 10: "окт", 11: "ноя", 12: "дек"}
MONTHS_EXCEL = {1: "янв.", 2: "февр.", 3: "март", 4: "апр.", 5: "май", 6: "июнь",
                7: "июль", 8: "авг.", 9: "сент.", 10: "окт.", 11: "нояб.", 12: "дек."}
MONTHS_NOM = {1: "январь", 2: "февраль", 3: "март", 4: "апрель", 5: "май", 6: "июнь",
              7: "июль", 8: "август", 9: "сентябрь", 10: "октябрь", 11: "ноябрь",
              12: "декабрь"}
CORE_KEYS = ["vol", "trim", "ex1", "ex2"]
TOP_CODES = ("ПР", "НЕПР", "У")
TOP_TITLES = {"ПР": "продовольственные", "НЕПР": "непродовольственные", "У": "услуги"}


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
    values: pd.DataFrame        # индексы м./м. (пред. месяц = 100): месяцы × категории
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
    _assign_sa_groups(meta, list(sa_values.columns))
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


def _assign_sa_groups(meta: pd.DataFrame, sa_cols: list) -> None:
    """
    Служебные SA-категории («Прочие ... (для SA)») нужны только для сезонного
    сглаживания: группы на листе «Данные» у них нет намеренно, поэтому в детальную
    корзину, расчеты продов/непродов/услуг и топ-5 они не попадают. SA-хедлайн при
    этом собирается из крупных категорий вместе с ними, поэтому для разложения на
    вклады каждая SA-категория относится к своей группе (ПР / НЕПР / У): по коду,
    а если кода нет — по блоку на листе SA (как в формулах «SA Продовольственые
    товары» и др., где «Прочие» замыкают блок своей группы).
    """
    tokens = meta["tokens"].to_dict()
    group = {}
    for c in sa_cols:
        found = [g for g in TOP_CODES if g in tokens[c]]
        group[c] = found[0] if len(found) == 1 else None
    for k, c in enumerate(sa_cols):
        if group[c] is not None:
            continue
        before = [group[x] for x in sa_cols[:k] if group[x]]
        after = [group[x] for x in sa_cols[k + 1:] if group[x]]
        g = before[-1] if before else (after[0] if after else None)
        if g is not None:
            group[c] = g
            tokens[c] = tokens[c] | {g}
    meta["tokens"] = pd.Series([tokens[i] for i in meta.index], index=meta.index, dtype=object)
    meta["sa_group"] = pd.Series([group.get(i) for i in meta.index], index=meta.index, dtype=object)
    meta["sa_service"] = pd.Series(
        [i in group and _is_na(meta.at[i, "group"]) for i in meta.index], index=meta.index)


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
    """Г./г. из цепочки м./м. (%): произведение 12 последовательных месячных индексов."""
    log_factor = np.log1p(mm.astype(float) / 100.0)
    return np.expm1(log_factor.rolling(12, min_periods=12).sum()) * 100.0


def saar_from_mm(mm):
    """SAAR: м./м. в годовом выражении."""
    return ((1.0 + mm / 100.0) ** 12 - 1.0) * 100.0


def saar_3m(mm):
    """SAAR 3м./3м.: прирост за последние 3 месяца в годовом выражении."""
    log_factor = np.log1p(mm.astype(float) / 100.0)
    return np.expm1(log_factor.rolling(3, min_periods=3).sum() * 4.0) * 100.0


def laspeyres_weights(inp: InflationInput, cols) -> pd.DataFrame:
    """
    Веса для агрегирования м./м. по схеме Росстата (цепной индекс Ласпейреса).

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
    """Взвешенное среднее м./м. по доступным в каждом месяце категориям."""
    w = ew.where(mm.notna() & (ew > 0))
    return (mm * w).sum(axis=1, min_count=1) / w.sum(axis=1, min_count=1)


def detailed_coverage(inp: InflationInput) -> pd.Series:
    """Доля веса (в %), покрытая детальными категориями (группа 0) с данными."""
    cols0 = inp.meta.index[inp.meta["group"].eq("0")]
    w = inp.weights[cols0].where(inp.values[cols0].notna())
    return w.sum(axis=1, min_count=1)


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


def ytd_table(mm: pd.Series) -> pd.DataFrame:
    """Таблица для листа: «Месяц» + столбцы-годы (накопленная с начала года, %)."""
    table = ytd_by_year(mm)
    table.columns = [str(c) for c in table.columns]
    table.insert(0, "Месяц", [MONTHS_SHORT[m] for m in table.index])
    return table.reset_index(drop=True)


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
    mm: pd.DataFrame            # м./м., %
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
        "trim": "усечение",
        "ex1": "исключение",
        "ex2": "исключение, без туризма",
    }


def weighted_trimmed_mean(x: np.ndarray, w: np.ndarray, low: float, high: float) -> tuple[float, np.ndarray]:
    """
    Взвешенное усеченное среднее: отрезаются доли low и high накопленного веса
    снизу и сверху распределения м./м. Категория на границе входит частично.
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
    Четыре меры базовой инфляции (м./м., %) и диагностика на последнюю дату.

    1) Без 20% наиболее волатильных: в каждом месяце t по окну из последних
       CORE_VOLATILITY_WINDOW месяцев (t включительно) считается дисперсия м./м.
       каждой категории; исключаются 20% категорий с наибольшей дисперсией.
       Категории без полного окна (новые) не ранжируются и остаются в корзине.
       Поэтому ряд начинается на (окно − 1) месяца позже начала данных.
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
    tokens = basis.tokens.to_numpy()
    p_latest = basis.mm.index.get_loc(latest) if latest in basis.mm.index else None

    for p in range(len(basis.mm)):
        x, w = values[p], weights[p]
        cur = ~np.isnan(x) & ~np.isnan(w) & (w > 0)
        if not cur.any():
            continue

        def wmean(mask):
            return float((x[mask] * w[mask]).sum() / w[mask].sum()) if mask.any() else np.nan

        # 1) Без наиболее волатильных.
        dropped, dropped_std, n_eligible = np.array([], dtype=int), np.array([]), 0
        if p >= window - 1:
            idx = np.flatnonzero(cur)
            hist = values[p - window + 1:p + 1, idx]
            full = ~np.isnan(hist).any(axis=0)
            eligible = idx[full]
            variance = hist[:, full].var(axis=0, ddof=1)
            n_drop = math.ceil(round(CORE_VOLATILITY_SHARE * len(eligible), 9))
            top = np.argsort(-variance, kind="mergesort")[:n_drop]
            dropped, dropped_std, n_eligible = eligible[top], np.sqrt(variance[top]), len(eligible)
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
            total_w = w[cur].sum()

            def excluded(keep):
                mask = cur & ~keep
                return [(n, ww / total_w * 100, t) for n, ww, t in zip(names[mask], w[mask], tokens[mask])]

            diag = {
                "vol_dropped": list(zip(names[dropped], dropped_std)),
                "vol_eligible": n_eligible,
                "trim_low": sorted([(n, v) for n, v, t in zip(cur_names, x[cur], trimmed_share)
                                    if t > 1e-9 and v <= median_x], key=lambda item: item[1]),
                "trim_high": sorted([(n, v) for n, v, t in zip(cur_names, x[cur], trimmed_share)
                                     if t > 1e-9 and v > median_x], key=lambda item: -item[1]),
                "ex1": excluded(keep1),
                "ex2": excluded(keep2),
                "n": int(cur.sum()),
            }

    result = pd.DataFrame(out, index=basis.mm.index, columns=CORE_KEYS)
    diag["codes_not_found"] = sorted(c for c in ex2 if not _has_any_code(basis.tokens, {c}).any())
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


def main_columns(kind: str) -> dict:
    """Названия столбцов листа 1 (используются и в таблице, и в графиках)."""
    labels = core_labels()
    return {
        "hl_nsa": f"Хедлайн nSA, {MM}, %",
        "hl_sa": f"Хедлайн SA, {MM}, %",
        "core_mm": {k: f"Базовая {kind}: {labels[k]}, {MM}, %" for k in CORE_KEYS},
        "core_mm_min": f"Базовая {kind}: мин., {MM}, %",
        "core_mm_max": f"Базовая {kind}: макс., {MM}, %",
        "hl_yoy": f"Хедлайн, {YY}, %",
        "core_yoy": {k: f"Базовая: {labels[k]}, {YY}, %" for k in CORE_KEYS},
        "core_yoy_min": f"Базовая: мин., {YY}, %",
        "core_yoy_max": f"Базовая: макс., {YY}, %",
        "hl_saar": "Хедлайн SAAR, %",
        "core_saar": {k: f"Базовая: {labels[k]}, SAAR, %" for k in CORE_KEYS},
        "hl_saar3": "Хедлайн SAAR 3м./3м., %",
        "band_mm": f"Для графика: ширина диапазона базовой {MM} (макс. − мин.), п.п.",
        "band_yoy": f"Для графика: ширина диапазона базовой {YY} (макс. − мин.), п.п.",
    }


def calc_main(inp, hl_mm, hl_sa_mm, core, basis) -> tuple[pd.DataFrame, dict]:
    c = main_columns(basis.kind)
    core_yoy = core.apply(yoy_from_mm)
    table = {c["hl_nsa"]: hl_mm, c["hl_sa"]: hl_sa_mm}
    for k in CORE_KEYS:
        table[c["core_mm"][k]] = core[k]
    table[c["core_mm_min"]] = core.min(axis=1)
    table[c["core_mm_max"]] = core.max(axis=1)
    table[c["hl_yoy"]] = yoy_from_mm(hl_mm)
    for k in CORE_KEYS:
        table[c["core_yoy"][k]] = core_yoy[k]
    table[c["core_yoy_min"]] = core_yoy.min(axis=1)
    table[c["core_yoy_max"]] = core_yoy.max(axis=1)
    table[c["hl_saar"]] = saar_from_mm(hl_sa_mm)
    if basis.kind == "SA":
        for k in CORE_KEYS:
            table[c["core_saar"][k]] = saar_from_mm(core[k])
    table[c["hl_saar3"]] = saar_3m(hl_sa_mm)
    table[c["band_mm"]] = core.max(axis=1) - core.min(axis=1)
    table[c["band_yoy"]] = core_yoy.max(axis=1) - core_yoy.min(axis=1)
    df = _cut(pd.DataFrame(table), inp.latest)
    df.insert(0, "Дата", df.index)
    return df.reset_index(drop=True), {"kind": basis.kind, "title": basis.title, "columns": c}


def quarter_gapped(df: pd.DataFrame) -> pd.DataFrame:
    """Помесячная таблица → блоки по кварталам с пустой строкой между ними."""
    rows = []
    for (year, quarter), block in df.groupby([df.index.year, df.index.quarter]):
        for i, (dt, values) in enumerate(block.iterrows()):
            rows.append({"Квартал": f"{quarter} кв. {year}" if i == 0 else "",
                         "Месяц": MONTHS_SHORT[dt.month], "Дата": dt, **values.to_dict()})
        rows.append({"Квартал": " ", "Месяц": "", "Дата": pd.NaT})
    return pd.DataFrame(rows[:-1], columns=["Квартал", "Месяц", "Дата", *df.columns])


ZERO_COLUMN = "Ноль (служебный ряд: линия оси на графике)"


def calc_contributions(inp, basis, hl_mm, hl_sa_mm) -> tuple[pd.DataFrame, dict]:
    assign, names = assign_components(inp, CONTRIBUTIONS, basis.cols)
    unassigned = assign.index[assign < 0]
    if len(unassigned):
        w_last = basis.ew.loc[inp.latest, unassigned]
        share = float(w_last.sum() / basis.ew.loc[inp.latest].sum() * 100)
        print(f"⚠️ Вклады: {len(unassigned)} кат. ({share:.1f}% веса) не попали ни в один компонент "
              f"и показаны отдельно как «Не распределено»: "
              f"{inp.meta.loc[unassigned[:8], 'name'].tolist()}{' …' if len(unassigned) > 8 else ''}. "
              f"Добавьте компонент {{\"остальное\": True}} или заполните разметку.")
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
        line_name, line = f"Хедлайн SA, {MM}, %", hl_sa_mm
    else:
        line_name, line = f"Хедлайн nSA, {MM}, %", hl_mm
    table = pd.DataFrame({line_name: line, **parts})
    residual = line - pd.DataFrame(parts).sum(axis=1, min_count=1)
    if residual.abs().max() > 0.005:
        table["Расхождение с хедлайном, п.п."] = residual
    table = _cut(table.loc[denominator.dropna().index.min():], inp.latest)
    gapped = quarter_gapped(table)
    gapped[ZERO_COLUMN] = 0.0
    info = {"line": line_name, "components": [c for c in table.columns if c != line_name],
            "basis": basis}
    return gapped, info


def calc_cumulative(inp, hl_mm) -> tuple[pd.DataFrame, dict]:
    table = ytd_table(hl_mm.loc[:inp.latest])
    current = str(inp.latest.year)
    previous = str(inp.latest.year - 1)
    pale = [str(inp.latest.year - k) for k in range(2, 2 + int(CUMULATIVE_PALE_YEARS))]
    pale = [y for y in pale if y in table.columns]
    return table, {"current": current, "previous": previous, "pale": pale}


def calc_split(inp, coverage) -> tuple[pd.DataFrame, dict]:
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
        f"{n1}, {MM}, %": s1, f"{n0}, {MM}, %": s0,
        f"{n1}, {YY}, %": yoy_from_mm(s1), f"{n0}, {YY}, %": yoy_from_mm(s0),
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
    return df.reset_index(drop=True), {"names": SPLIT_NAMES}


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
    disp_name = (f"Разброс {MM} по корзине (станд. откл.), п.п." if kind == "std"
                 else f"Разброс {MM} по корзине (межкварт. размах), п.п.")
    infl_name = f"Хедлайн SA, {MM}, %" if basis.kind == "SA" else f"Хедлайн nSA, {MM}, %"
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
    """Топ-N по м./м., г./г. и вкладу в хедлайн на последнюю дату (вклады — по весам Ласпейреса)."""
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
        metrics = {f"{MM}, %": mm.loc[latest, group_cols], f"{YY}, %": yoy.loc[group_cols],
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
    table = _cut(pd.DataFrame({f"{MM}, %": mm, f"{YY}, %": yoy_from_mm(mm)}), inp.latest)
    if table.empty:
        print(f"⚠️ У категории «{name}» нет данных.")
        return None
    table.insert(0, "Дата", table.index)
    ytd = ytd_table(mm)
    if str(inp.latest.year) not in ytd.columns:
        print(f"ℹ️ «{name}»: накопленная с начала {inp.latest.year} г. не считается — "
              f"нет данных за январь или ряд прерывается.")
    return {"name": str(inp.meta.loc[cid, "name"]), "table": table.reset_index(drop=True),
            "ytd": ytd, "current_year": str(inp.latest.year)}


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
    split, split_info = calc_split(inp, coverage)
    cumulative, cumulative_info = calc_cumulative(inp, hl_mm)
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


def _wrap(items: list[str], indent: str = "       ") -> str:
    return textwrap.fill(", ".join(items), width=110, initial_indent=indent,
                         subsequent_indent=indent)


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
        print(f"   Допуск ±{tol:.2f} п.п. {MM} Покрытие весом на последнюю дату: "
              f"{coverage.get(latest, np.nan):.1f}%.")
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

    # 2) SA-хедлайн и вклады.
    if sa_basis is not None:
        print("-" * 78)
        rebuilt_sa = weighted_mm(sa_basis.mm, sa_basis.ew)
        diff = (rebuilt_sa - hl_sa_mm).loc[:latest].abs().max()
        status = "✅" if diff < 0.005 else "⚠️"
        print(f"2) {status} SA-хедлайн, пересчитанный из крупных SA-категорий, vs столбец "
              f"«SA Все товары и услуги»: макс. расхождение {diff:.1e} п.п.")
        # Вклады групп по крупным категориям vs ваши столбцы «SA Продовольственые товары»
        # и т.д. Молчит, пока все сходится; напишет, если, например, переставить столбцы
        # на листе SA и «Прочие» окажутся не в своем блоке.
        pairs = [("ПР", "food", "sa_food"), ("НЕПР", "nonfood", "sa_nonfood"), ("У", "services", "sa_services")]
        if all(inp.cols.get(a) is not None and inp.cols.get(b) is not None for _, a, b in pairs):
            w = sa_basis.ew.where(sa_basis.mm.notna())
            groups = meta.loc[sa_basis.cols, "sa_group"]
            agg_ew = laspeyres_weights(inp, [inp.cols[a] for _, a, _ in pairs])
            worst = 0.0
            for g, a, b in pairs:
                own = groups.index[groups.eq(g)]
                by_categories = (sa_basis.mm[own] * w[own]).sum(axis=1) / w.sum(axis=1)
                by_aggregate = (agg_ew[inp.cols[a]] * (inp.values[inp.cols[b]] - 100.0)
                                / agg_ew.sum(axis=1))
                worst = max(worst, float((by_categories - by_aggregate).loc[:latest].abs().max()))
            if worst >= 0.005:
                print(f"   ⚠️ Вклады продовольствия, непрода и услуг расходятся с расчетом по столбцам "
                      f"«SA ...» на {worst:.3f} п.п.: проверьте коды и порядок столбцов листа SA.")
        weight_sum = inp.weights[sa_basis.cols].where(sa_basis.mm.notna()).sum(axis=1)
        bad_years = sorted({d.year for d, v in weight_sum.loc[:latest].items() if abs(v - 100) > 0.05})
        if bad_years:
            print(f"   ⚠️ Сумма весов крупных SA-категорий ≠ 100 в годах: {bad_years}")
        sa_meta = meta.loc[[c for c in sa_basis.cols if not meta.at[c, "sa_service"]]]
        blank = {}
        for key, label in [("code", "Код"), *[(k, v) for k, v in inp.flags.items() if k != "sa"]]:
            for i in sa_meta.index[sa_meta[key].isna()]:
                blank.setdefault(sa_meta.at[i, "name"], []).append(label)
        for name, fields in blank.items():
            print(f"   ⚠️ «{name}»: не заполнено — {', '.join(fields)}")
        yoy_gap = (yoy_from_mm(hl_sa_mm) - yoy_from_mm(hl_mm)).get(latest, np.nan)
        if abs(yoy_gap) > 0.3:
            print(f"   ℹ️ SA-хедлайн {YY} отличается от официального {YY} на {yoy_gap:+.2f} п.п. "
                  f"(так бывает при переносе сроков индексации тарифов).")

    # 3) Базовая инфляция.
    print("-" * 78)
    labels = core_labels()
    print(f"3) Базовая инфляция: {core_basis.title}, {latest:%m.%Y}")
    if core_diag.get("codes_not_found"):
        print(f"   ⚠️ Коды исключения не найдены ни у одной категории: {core_diag['codes_not_found']} "
              f"— эти товары/услуги НЕ исключаются.")
    if core_diag.get("n"):
        dropped = core_diag["vol_dropped"]
        print(f"   • {labels['vol'].capitalize()}: исключено {len(dropped)} из "
              f"{core_diag['vol_eligible']} кат. (в скобках — станд. откл. {MM} за "
              f"{CORE_VOLATILITY_WINDOW} мес., п.п.):")
        print(_wrap([f"{n} ({s:.2f})" for n, s in dropped]))
        low = [f"{n} ({v:+.2f})" for n, v in core_diag["trim_low"]]
        high = [f"{n} ({v:+.2f})" for n, v in core_diag["trim_high"]]
        if core_basis.kind == "SA":
            print(f"   • Усечение — срезаны снизу ({MM}, %):")
            print(_wrap(low))
            print("     сверху:")
            print(_wrap(high))
        else:
            print(f"   • Усечение: срезано снизу {len(low)} кат., сверху {len(high)} кат.")
        for key, title in [("ex1", labels["ex1"]), ("ex2", labels["ex2"])]:
            items = core_diag[key]
            share = sum(s for _, s, _ in items)
            print(f"   • {title.capitalize()}: исключено {share:.1f}% веса корзины:")
            if core_basis.kind == "SA":
                print(_wrap([f"{n} ({s:.1f}%)" for n, s, _ in items]))
            else:
                codes = CORE_EXCLUDE_CODES + (CORE_EXCLUDE_TOURISM_CODES if key == "ex2" else [])
                summary = []
                for code in codes:
                    sel = [(n, s) for n, s, t in items if code.upper() in t]
                    if sel:
                        summary.append(f"{code}: {len(sel)} кат., {sum(s for _, s in sel):.1f}%")
                print(_wrap(summary) + "  (категория с несколькими кодами учтена в каждом)")
    print(line)


# =============================================================================
# 7. EXCEL: ТАБЛИЦЫ + РЕДАКТИРУЕМЫЕ ГРАФИКИ
# =============================================================================


class _Styles:
    def __init__(self, wb):
        text = {"font_name": FONT, "font_size": FONT_SIZE, "font_color": AXIS_COLOR}
        self.header = wb.add_format({
            "font_name": FONT, "font_size": 9, "font_color": "#FFFFFF",
            "bg_color": BRAND["navy"], "text_wrap": True, "align": "center",
            "valign": "vcenter", "border": 1, "border_color": "#FFFFFF",
        })
        self.date = wb.add_format({**text, "num_format": "dd.mm.yyyy", "align": "left"})
        self.num = wb.add_format({**text, "num_format": "0.00"})
        self.text = wb.add_format(text)
        self.title = wb.add_format({"font_name": FONT, "font_size": 12, "bold": True})
        self.section = wb.add_format({"font_name": FONT, "font_size": FONT_SIZE,
                                      "font_color": "#FFFFFF", "bg_color": BRAND["navy"]})
        self.subheader = wb.add_format({"font_name": FONT, "font_size": 9,
                                        "bg_color": BRAND["light_grey"], "align": "center",
                                        "text_wrap": True, "valign": "vcenter"})
        self.note = wb.add_format({**text, "italic": True, "font_size": 9,
                                   "font_color": BRAND["grey_blue"]})


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


AXIS_LINE = {"color": AXIS_COLOR, "width": 0.75}


def _font(size=FONT_SIZE, bold=False, color=AXIS_COLOR, **extra) -> dict:
    return {"name": FONT, "size": size, "bold": bold, "color": color, **extra}


def _style_chart(chart, title: str, *, date_axis: bool = True, legend: bool = True,
                 size=CHART_SIZE, y_at_right: bool = True, x_line: bool = True,
                 y_format: str = "0.0", delete_from_legend: list | None = None) -> None:
    """Общее оформление: заголовок слева жирным, легенда снизу, черные оси, ось Y справа."""
    chart.set_title({"name": title, "name_font": _font(TITLE_SIZE, bold=True), "overlay": False,
                     "layout": {"x": 0.01, "y": 0.02}})
    if legend:
        options = {"position": "bottom", "font": _font()}
        if delete_from_legend:
            options["delete_series"] = delete_from_legend
        chart.set_legend(options)
    else:
        chart.set_legend({"none": True})
    x_axis = {"num_font": _font(), "line": AXIS_LINE if x_line else {"none": True},
              "label_position": "low", "major_gridlines": {"visible": False},
              "major_tick_mark": "outside" if date_axis else "none",
              "position_axis": "between"}
    if y_at_right:
        x_axis["crossing"] = "max"      # ось Y пересекает ось X в максимальном значении
    if date_axis:
        x_axis.update({"date_axis": True, "num_format": DATE_AXIS_FORMAT,
                       "major_unit": 1, "major_unit_type": "months",
                       "base_unit": 1, "base_unit_type": "months",
                       "num_font": _font(rotation=-90)})
    chart.set_x_axis(x_axis)
    chart.set_y_axis({"num_font": _font(), "num_format": y_format, "line": AXIS_LINE,
                      "major_tick_mark": "outside",
                      "major_gridlines": {"visible": CHART_GRIDLINES,
                                          "line": {"color": BRAND["light_grey"], "width": 0.5}}})
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


def _chart_height_rows(height_px: int) -> int:
    return int(math.ceil(height_px / 20.0)) + 2


def _sheet_main(wb, st, report: ReportData) -> None:
    name = "Хедлайн и базовая"
    ws = wb.add_worksheet(name)
    df, info = report.main, report.main_info
    c = info["columns"]
    _write_table(ws, st, df)
    ws.freeze_panes(1, 1)
    rows = _chart_rows(df["Дата"], report.chart_start)
    if rows is None:
        return
    r1, r2 = rows
    col = {name_: j for j, name_ in enumerate(df.columns)}
    labels = {k: v[0].upper() + v[1:] for k, v in core_labels().items()}
    cats = [name, r1, 0, r2, 0]
    chart_col = len(df.columns) + 1

    def ref(column):
        return [name, r1, col[column], r2, col[column]]

    def core_band(chart, low_column, width_column, label):
        # Диапазон базовой инфляции: невидимая «подложка» до минимума + заливка
        # высотой (макс. − мин.). Подложка удалена из легенды.
        chart.add_series({"name": "мин.", "categories": cats, "values": ref(low_column),
                          "fill": {"none": True}, "line": {"none": True}})
        chart.add_series({"name": label, "categories": cats, "values": ref(width_column),
                          "fill": {"color": CORE_BAND_COLOR, "transparency": CORE_BAND_TRANSPARENCY},
                          "line": {"none": True}})

    kind = info["kind"]
    # (1) м./м.: диапазон базовой, SA-хедлайн ярко, nSA-хедлайн бледно.
    area = wb.add_chart({"type": "area", "subtype": "stacked"})
    core_band(area, c["core_mm_min"], c["band_mm"], f"Базовая {kind}: диапазон 4 методов")
    lines = wb.add_chart({"type": "line"})
    lines.add_series({"name": "Хедлайн nSA", "categories": cats, "values": ref(c["hl_nsa"]),
                      "line": _line(HEADLINE_NSA_COLOR, 1.5)})
    lines.add_series({"name": "Хедлайн SA", "categories": cats, "values": ref(c["hl_sa"]),
                      "line": _line(HEADLINE_COLOR, 2.25)})
    area.combine(lines)
    _style_chart(area, f"Инфляция {MM}: хедлайн и базовая ({kind}), %", delete_from_legend=[0])
    ws.insert_chart(1, chart_col, area)

    # (2) г./г.: диапазон базовой и хедлайн линией; SAAR — точками разной формы.
    # Пустые ячейки Excel рисует у области как ноль, поэтому график начинается с
    # первого месяца, где есть г./г. базовой (важно для варианта "nSA").
    first_band = df.loc[df[c["core_yoy_min"]].notna(), "Дата"].min()
    if not pd.isna(first_band) and first_band > report.chart_start:
        rows_yoy = _chart_rows(df["Дата"], first_band)
        cats = [name, rows_yoy[0], 0, rows_yoy[1], 0]
        r1 = rows_yoy[0]
    area = wb.add_chart({"type": "area", "subtype": "stacked"})
    core_band(area, c["core_yoy_min"], c["band_yoy"], f"Базовая: диапазон 4 методов, {YY}")
    lines = wb.add_chart({"type": "line"})
    lines.add_series({"name": f"Хедлайн, {YY}", "categories": cats, "values": ref(c["hl_yoy"]),
                      "line": _line(HEADLINE_COLOR, 2.25)})
    lines.add_series({"name": "Хедлайн, SAAR", "categories": cats, "values": ref(c["hl_saar"]),
                      "line": {"none": True}, "marker": _marker("circle", HEADLINE_COLOR, 6)})
    if kind == "SA":
        for k in CORE_KEYS:
            color, marker = CORE_SAAR_MARKERS[k]
            lines.add_series({"name": f"{labels[k]}, SAAR", "categories": cats,
                              "values": ref(c["core_saar"][k]),
                              "line": {"none": True}, "marker": _marker(marker, color, 5)})
    area.combine(lines)
    size = (CHART_SIZE[0], CHART_SIZE[1] + 40)
    _style_chart(area, f"Инфляция {YY} и SAAR: хедлайн и базовая, %", size=size,
                 delete_from_legend=[0])
    ws.insert_chart(1 + _chart_height_rows(CHART_SIZE[1]), chart_col, area)


def _sheet_contributions(wb, st, report: ReportData) -> None:
    df, info = report.contrib, report.contrib_info
    name = "Вклады"
    ws = wb.add_worksheet(name)
    _write_table(ws, st, df, widths={c: 14 for c in df.columns})
    ws.freeze_panes(1, 3)
    zero_j = list(df.columns).index(ZERO_COLUMN)
    ws.set_column(zero_j, zero_j, 14, None, {"hidden": True})
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
            series["gap"] = COLUMN_GAP
        ch.add_series(series)
    lines = wb.add_chart({"type": "line"})
    # Линия нуля вместо оси X: у двухуровневой оси Excel рисует между кварталами
    # вертикальные разделители цветом линии оси, поэтому сама ось скрыта.
    lines.add_series({"name": "0", "categories": cat_ref, "categories_data": cats_cache,
                      "values": [name, r1, zero_j, r2, zero_j], "line": _line(AXIS_COLOR, 0.75)})
    lines.add_series({"name": info["line"].replace(", %", ""),
                      "categories": cat_ref, "categories_data": cats_cache,
                      "values": [name, r1, col[info["line"]], r2, col[info["line"]]],
                      "line": _line(AXIS_COLOR, 1.75)})
    ch.combine(lines)
    ch.show_blanks_as("span")
    ch.show_hidden_data()
    title = (f"Вклады в SA-хедлайн {MM}, п.п." if info["basis"].kind == "SA"
             else f"Вклады в хедлайн {MM} (nSA, детальные категории), п.п.")
    _style_chart(ch, title, date_axis=False, x_line=False,
                 delete_from_legend=[len(info["components"])])
    ws.insert_chart(1, len(df.columns) + 1, ch)


def _past_years_label(years: list) -> str:
    """Подпись в легенде для всех бледно-серых линий сразу: «Прошлые годы (2002–2024)»."""
    first, last = min(years), max(years)
    return f"Прошлые годы ({first})" if first == last else f"Прошлые годы ({first}–{last})"


def _ytd_chart(wb, sheet: str, ytd: pd.DataFrame, col0: int, title: str, current: str,
               previous: str, pale: list | None) -> object | None:
    """
    Накопленная с начала года: текущий год — розовым, прошлый — темно-синим.
    pale — список лет для бледных линий разного цвета; None — все остальные годы
    светло-серым, в легенде — одной строкой «Прошлые годы (…)».
    """
    years = list(ytd.columns[1:])
    if not years:
        return None
    if pale is None:
        older = [y for y in years if y not in (current, previous)]
        styles = {y: _line(BRAND["light_grey"], 1.0) for y in older}
        hidden_years = set(older)
    else:
        older = [y for y in pale if y in years]
        styles = {y: _line(PALE_YEAR_COLORS[min(k, len(PALE_YEAR_COLORS) - 1)], 1.5)
                  for k, y in enumerate(older)}
        hidden_years = set()
    styles[previous] = _line(PREVIOUS_YEAR_COLOR, 2.25)
    styles[current] = _line(CURRENT_YEAR_COLOR, 2.75)
    order = list(reversed(older)) + [y for y in (previous, current) if y in years]
    ch = wb.add_chart({"type": "line"})
    if hidden_years:
        # Строка легенды для серых линий — отдельная пустая серия (ссылается на пустой
        # столбец сразу справа от таблицы): так у каждой серой линии остается ее год
        # в подсказке Excel, а в легенде они не перечисляются по одной.
        gap = col0 + len(ytd.columns)
        ch.add_series({"name": _past_years_label(sorted(hidden_years)),
                       "categories": [sheet, 1, col0, 12, col0],
                       "values": [sheet, 1, gap, 12, gap], "line": _line(BRAND["light_grey"], 1.0)})
    shift = 1 if hidden_years else 0
    hidden = []
    for i, year in enumerate(order):
        j = col0 + list(ytd.columns).index(year)
        ch.add_series({"name": year, "categories": [sheet, 1, col0, 12, col0],
                       "values": [sheet, 1, j, 12, j], "line": styles[year]})
        if year in hidden_years:
            hidden.append(i + shift)
    _style_chart(ch, title, date_axis=False, delete_from_legend=hidden or None)
    return ch


def _sheet_cumulative(wb, st, report: ReportData) -> None:
    df, info = report.cumulative, report.cumulative_info
    name = "Накопленная"
    ws = wb.add_worksheet(name)
    _write_table(ws, st, df, widths={c: 8 for c in df.columns})
    ws.freeze_panes(1, 1)
    ch = _ytd_chart(wb, name, df, 0, "Накопленная с начала года инфляция, %",
                    info["current"], info["previous"], info["pale"])
    if ch is not None:
        ws.insert_chart(1, len(df.columns) + 1, ch)


def _sheet_split(wb, st, report: ReportData) -> None:
    df, info = report.split, report.split_info
    n1, n0 = info["names"]
    name = f"{n1} и {n0.lower()}"[:31]
    ws = wb.add_worksheet(name)
    _write_table(ws, st, df, widths={c: 15 for c in df.columns})
    ws.freeze_panes(1, 1)
    row = 1
    for measure in (MM, YY):
        # Г./г. появляется через 12 месяцев после начала детальных данных:
        # график начинается с первого месяца, где есть значения.
        shown = df[[f"{n1}, {measure}, %", f"{n0}, {measure}, %"]].notna().any(axis=1)
        first_valid = df.loc[shown, "Дата"].min()
        start = report.chart_start if pd.isna(first_valid) else max(report.chart_start, first_valid)
        rows = _chart_rows(df["Дата"], start)
        if rows is None:
            continue
        r1, r2 = rows
        ch = wb.add_chart({"type": "line"})
        for series_name, color in [(n1, BRAND["navy"]), (n0, BRAND["magenta"])]:
            j = list(df.columns).index(f"{series_name}, {measure}, %")
            ch.add_series({"name": series_name, "categories": [name, r1, 0, r2, 0],
                           "values": [name, r1, j, r2, j], "line": _line(color, 2.25)})
        _style_chart(ch, f"{n1} и {n0.lower()} инфляция, {measure}, %")
        ws.insert_chart(row, len(df.columns) + 1, ch)
        row += _chart_height_rows(CHART_SIZE[1])


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
    sizes = (9, 7, 6)
    ch = wb.add_chart({"type": "scatter"})
    # Сначала прошлые годы, чтобы последние точки рисовались поверх.
    for t in (2, 1, 0):
        tier_name = info["tiers"][t]
        if part[tier_name].notna().sum() == 0:
            continue  # например, в январе нет точек «ранее в текущем году»
        custom = [{"value": lab, "font": _font(9, color=MATRIX_LABEL_COLORS[t])}
                  if not _is_na(v) else {"delete": True}
                  for lab, v in zip(part["Метка"], part[tier_name])]
        ch.add_series({"name": tier_name,
                       "categories": [name, r1, col[info["x"]], r2, col[info["x"]]],
                       "values": [name, r1, col[tier_name], r2, col[tier_name]],
                       "marker": _marker("circle", MATRIX_COLORS[t], sizes[t]),
                       "data_labels": {"value": True, "custom": custom, "position": "right"}})
    ch.set_title({"name": "Инфляция и разброс изменений цен по корзине",
                  "name_font": _font(TITLE_SIZE, bold=True), "overlay": False,
                  "layout": {"x": 0.01, "y": 0.02}})
    ch.set_legend({"position": "bottom", "font": _font()})
    # Оси пересекаются на границах «высокая/низкая» и «однородная/неоднородная»:
    # получается решетка из четырех квадрантов, подписи осей — по краям.
    ch.set_x_axis({"name": info["x"], "name_font": _font(), "num_font": _font(),
                   "num_format": "0.0", "line": AXIS_LINE, "label_position": "low",
                   "crossing": round(info["x_threshold"], 3),
                   "major_gridlines": {"visible": False}})
    ch.set_y_axis({"name": info["y"], "name_font": _font(), "num_font": _font(),
                   "num_format": "0.0", "line": AXIS_LINE, "label_position": "low",
                   "crossing": round(info["y_threshold"], 3),
                   "major_gridlines": {"visible": False}})
    ch.set_chartarea({"border": {"none": True}, "fill": {"color": CHART_BACKGROUND}})
    ch.set_plotarea({"border": {"none": True}, "fill": {"none": True}})
    ch.set_size({"width": 760, "height": 600})
    ws.insert_chart(1, len(df.columns) + 1, ch)
    ws.write(_chart_height_rows(600) + 1, len(df.columns) + 1,
             f"Горизонтальная ось пересекает вертикальную на уровне {info['y_threshold']:.2f}% {MM} "
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
        for metric in [f"{MM}, %", f"{YY}, %", "Вклад в хедлайн, п.п."]:
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

    row = 1
    rows = _chart_rows(table["Дата"], report.chart_start)
    if rows is not None:
        r1, r2 = rows
        # Две шкалы: м./м. — левая, г./г. — правая (основная, «в максимальном значении»).
        ch = wb.add_chart({"type": "line"})
        ch.add_series({"name": f"{MM} (левая шкала)", "categories": [name, r1, 0, r2, 0],
                       "values": [name, r1, 1, r2, 1], "line": _line(BRAND["magenta"], 1.5)})
        ch.add_series({"name": f"{YY} (правая шкала)", "categories": [name, r1, 0, r2, 0],
                       "values": [name, r1, 2, r2, 2], "line": _line(BRAND["navy"], 2.25),
                       "y2_axis": True})
        _style_chart(ch, f"{cat['name']}: {MM} и {YY}, %", y_at_right=False)
        ch.set_y2_axis({"num_font": _font(), "num_format": "0.0", "line": AXIS_LINE,
                        "major_tick_mark": "outside", "major_gridlines": {"visible": False}})
        ws.insert_chart(row, chart_col, ch)
        row += _chart_height_rows(CHART_SIZE[1])

    # Накопленная с начала года: все прошлые годы бледно-серым (в легенде — «Прошлые годы (…)»).
    ch = _ytd_chart(wb, name, ytd, ytd_col, f"{cat['name']}: накопленная с начала года инфляция, %",
                    cat["current_year"], str(int(cat["current_year"]) - 1), None)
    if ch is not None:
        ws.insert_chart(row, chart_col, ch)


def export_excel(report: ReportData, path: Path) -> None:
    import xlsxwriter
    from xlsxwriter.exceptions import FileCreateError

    path = Path(path)
    wb = xlsxwriter.Workbook(str(path))
    wb.formats[0].set_font_name(FONT)
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

    style = {
        "font.family": "sans-serif",
        "font.sans-serif": [FONT, "Liberation Sans", "DejaVu Sans"],
        "font.size": FONT_SIZE, "axes.edgecolor": AXIS_COLOR, "axes.linewidth": 0.75,
        "xtick.color": AXIS_COLOR, "ytick.color": AXIS_COLOR, "legend.frameon": False,
        "axes.unicode_minus": False,
    }
    comma = FuncFormatter(lambda v, _: f"{v:.1f}".replace(".", ","))
    markers = {"circle": "o", "square": "s", "diamond": "D", "triangle": "^"}
    alpha = 1.0 - CORE_BAND_TRANSPARENCY / 100.0
    start = report.chart_start

    def new_axes(height=6.2, y_right=True):
        fig, ax = plt.subplots(figsize=(11.0, height))
        ax.spines["top"].set_visible(False)
        if y_right:
            ax.spines["left"].set_visible(False)
            ax.yaxis.tick_right()
        else:
            ax.spines["right"].set_visible(False)
        ax.yaxis.set_major_formatter(comma)
        return fig, ax

    def zero_line(ax):
        low, high = ax.get_ylim()
        if low < 0 < high:
            ax.spines["bottom"].set_visible(False)
            ax.axhline(0, color=AXIS_COLOR, lw=0.75, zorder=1)

    def month_axis(ax, dates):
        dates = list(pd.to_datetime(pd.Series(dates)).dropna())
        if not dates:
            return
        ax.set_xticks(dates)
        ax.set_xticklabels([f"{MONTHS_EXCEL[d.month]} {d.year % 100:02d}" for d in dates], rotation=90)
        ax.set_xlim(dates[0] - pd.Timedelta(days=16), dates[-1] + pd.Timedelta(days=16))

    def finish(fig, ax, title, handles=None, ncol=3):
        if handles is None:
            handles, labels = ax.get_legend_handles_labels()
        else:
            labels = [h.get_label() for h in handles]
        n_rows = math.ceil(len(handles) / ncol) if handles else 0
        if handles:
            fig.legend(handles, labels, loc="lower center", ncol=ncol, frameon=False,
                       bbox_to_anchor=(0.5, 0.005))
        fig.text(0.01, 0.985, title, ha="left", va="top", fontsize=TITLE_SIZE, fontweight="bold")
        fig.tight_layout(rect=(0, 0.01 + 0.042 * n_rows, 1, 0.94))
        pdf.savefig(fig)
        plt.close(fig)

    with matplotlib.rc_context(style), PdfPages(path) as pdf:
        # Лист 1: м./м. и г./г. + SAAR.
        c, kind = report.main_info["columns"], report.main_info["kind"]
        labels = {k: v[0].upper() + v[1:] for k, v in core_labels().items()}
        d = report.main[report.main["Дата"] >= start]
        fig, ax = new_axes()
        ax.fill_between(d["Дата"], d[c["core_mm_min"]], d[c["core_mm_max"]], color=CORE_BAND_COLOR,
                        alpha=alpha, lw=0, label=f"Базовая {kind}: диапазон 4 методов")
        ax.plot(d["Дата"], d[c["hl_nsa"]], color=HEADLINE_NSA_COLOR, lw=1.5, label="Хедлайн nSA")
        ax.plot(d["Дата"], d[c["hl_sa"]], color=HEADLINE_COLOR, lw=2.25, label="Хедлайн SA")
        month_axis(ax, d["Дата"])
        zero_line(ax)
        finish(fig, ax, f"Инфляция {MM}: хедлайн и базовая ({kind}), %")

        fig, ax = new_axes(6.6)
        ax.fill_between(d["Дата"], d[c["core_yoy_min"]], d[c["core_yoy_max"]], color=CORE_BAND_COLOR,
                        alpha=alpha, lw=0, label=f"Базовая: диапазон 4 методов, {YY}")
        ax.plot(d["Дата"], d[c["hl_yoy"]], color=HEADLINE_COLOR, lw=2.25, label=f"Хедлайн, {YY}")
        ax.scatter(d["Дата"], d[c["hl_saar"]], color=HEADLINE_COLOR, s=30, zorder=3, label="Хедлайн, SAAR")
        if kind == "SA":
            for k in CORE_KEYS:
                color, marker = CORE_SAAR_MARKERS[k]
                ax.scatter(d["Дата"], d[c["core_saar"][k]], color=color, s=22, marker=markers[marker],
                           zorder=3, label=f"{labels[k]}, SAAR")
        month_axis(ax, d["Дата"])
        zero_line(ax)
        finish(fig, ax, f"Инфляция {YY} и SAAR: хедлайн и базовая, %")

        # Лист 2: вклады, месяцы по кварталам.
        cdf, cinfo = report.contrib, report.contrib_info
        q_start = pd.Timestamp(start.year, 3 * ((start.month - 1) // 3) + 1, 1)
        first = cdf.index[cdf["Дата"].ge(q_start).fillna(False)]
        if len(first):
            part = cdf.loc[first[0]:].reset_index(drop=True)
            x = np.arange(len(part))
            width = 1.0 / (1.0 + COLUMN_GAP / 100.0)
            fig, ax = new_axes()
            pos, neg = np.zeros(len(part)), np.zeros(len(part))
            for i, component in enumerate(cinfo["components"]):
                vals = part[component].astype(float).fillna(0).to_numpy()
                color = PALETTE_SIMPLE[i % len(PALETTE_SIMPLE)]
                if component.startswith(("Вклад: Не распределено", "Расхождение")):
                    color = BRAND["light_grey"]
                bottom = np.where(vals >= 0, pos, neg)
                ax.bar(x, vals, bottom=bottom, color=color, width=width,
                       label=component.replace("Вклад: ", "").replace(", п.п.", ""))
                pos += np.where(vals > 0, vals, 0)
                neg += np.where(vals < 0, vals, 0)
            ax.spines["bottom"].set_visible(False)
            ax.axhline(0, color=AXIS_COLOR, lw=0.75, zorder=1)
            total = part[cinfo["line"]].astype(float)
            ok = total.notna().to_numpy()
            line, = ax.plot(x[ok], total[ok], color=AXIS_COLOR, lw=1.75,
                            label=cinfo["line"].replace(", %", ""))
            month_rows = part["Месяц"].fillna("").astype(str).str.len().to_numpy() > 0
            ax.set_xticks(x[month_rows])
            ax.set_xticklabels(part.loc[month_rows, "Месяц"], fontsize=8)
            ax.tick_params(axis="x", length=0)
            ax.set_xlim(-0.6, len(part) - 0.4)
            for i, q in enumerate(part["Квартал"].fillna("")):
                if str(q).strip():
                    n_months = int(month_rows[i:i + 3].sum())
                    ax.annotate(q, xy=(i + (n_months - 1) / 2, 0), xycoords=("data", "axes fraction"),
                                xytext=(0, -22), textcoords="offset points", ha="center", fontsize=9)
            bars = [h for h in ax.get_legend_handles_labels()[0] if h is not line]
            title = (f"Вклады в SA-хедлайн {MM}, п.п." if cinfo["basis"].kind == "SA"
                     else f"Вклады в хедлайн {MM} (nSA, детальные категории), п.п.")
            finish(fig, ax, title, bars + [line], ncol=4)

        # Лист 3: накопленная с начала года (хедлайн).
        def ytd_figure(ytd, title, current, previous, pale):
            years = list(ytd.columns[1:])
            if not years:
                return
            fig, ax = new_axes()
            months = range(12)
            if pale is None:
                older = [y for y in years if y not in (current, previous)]
                for k, y in enumerate(older):
                    ax.plot(months, ytd[y], color=BRAND["light_grey"], lw=1.0,
                            label=_past_years_label(older) if k == 0 else None)
            else:
                for k, y in reversed(list(enumerate(pale))):
                    if y in years:
                        ax.plot(months, ytd[y], color=PALE_YEAR_COLORS[min(k, len(PALE_YEAR_COLORS) - 1)],
                                lw=1.5, label=y)
            if previous in years:
                ax.plot(months, ytd[previous], color=PREVIOUS_YEAR_COLOR, lw=2.25, label=previous)
            if current in years:
                ax.plot(months, ytd[current], color=CURRENT_YEAR_COLOR, lw=2.75, label=current)
            ax.set_xticks(list(months))
            ax.set_xticklabels(ytd["Месяц"])
            ax.set_xlim(-0.5, 11.5)
            zero_line(ax)
            finish(fig, ax, title, ncol=7)

        info = report.cumulative_info
        ytd_figure(report.cumulative, "Накопленная с начала года инфляция, %",
                   info["current"], info["previous"], info["pale"])

        # Лист 4: монетарная / немонетарная, м./м. и г./г.
        sp = report.split
        n1, n0 = report.split_info["names"]
        for measure in (MM, YY):
            shown = sp[[f"{n1}, {measure}, %", f"{n0}, {measure}, %"]].notna().any(axis=1)
            split_start = max(start, sp.loc[shown, "Дата"].min()) if shown.any() else start
            d = sp[sp["Дата"] >= split_start]
            fig, ax = new_axes()
            ax.plot(d["Дата"], d[f"{n1}, {measure}, %"], color=BRAND["navy"], lw=2.25, label=n1)
            ax.plot(d["Дата"], d[f"{n0}, {measure}, %"], color=BRAND["magenta"], lw=2.25, label=n0)
            month_axis(ax, d["Дата"])
            zero_line(ax)
            finish(fig, ax, f"{n1} и {n0.lower()} инфляция, {measure}, %")

        # Лист 5: матрица.
        m, minfo = report.matrix, report.matrix_info
        d = m[m["Дата"] >= minfo["window_start"]]
        fig, ax = plt.subplots(figsize=(10.0, 7.6))
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.axhline(minfo["y_threshold"], color=AXIS_COLOR, lw=0.75, zorder=1)
        ax.axvline(minfo["x_threshold"], color=AXIS_COLOR, lw=0.75, zorder=1)
        sizes = (70, 45, 30)
        for t in (2, 1, 0):
            tier = minfo["tiers"][t]
            sel = d[tier].notna()
            ax.scatter(d.loc[sel, minfo["x"]], d.loc[sel, tier], color=MATRIX_COLORS[t], s=sizes[t],
                       label=tier, zorder=3 + (2 - t))
            for _, r in d[sel].iterrows():
                ax.annotate(r["Метка"], (r[minfo["x"]], r[tier]), xytext=(5, 0),
                            textcoords="offset points", fontsize=8, color=MATRIX_LABEL_COLORS[t],
                            va="center")
        ax.set_xlabel(minfo["x"])
        ax.set_ylabel(minfo["y"])
        ax.xaxis.set_major_formatter(comma)
        ax.yaxis.set_major_formatter(comma)
        handles, labels_ = ax.get_legend_handles_labels()
        order = [labels_.index(t) for t in minfo["tiers"] if t in labels_]
        finish(fig, ax, "Инфляция и разброс изменений цен по корзине", [handles[i] for i in order])

        # Листы 7+: категории.
        for cat in report.categories:
            t = cat["table"]
            d = t[t["Дата"] >= start]
            fig, ax = new_axes(y_right=False)
            ax.plot(d["Дата"], d[f"{MM}, %"], color=BRAND["magenta"], lw=1.5, label=f"{MM} (левая шкала)")
            ax2 = ax.twinx()
            ax2.plot(d["Дата"], d[f"{YY}, %"], color=BRAND["navy"], lw=2.25, label=f"{YY} (правая шкала)")
            ax2.spines["top"].set_visible(False)
            ax2.spines["left"].set_visible(False)
            ax2.yaxis.set_major_formatter(comma)
            month_axis(ax, d["Дата"])
            finish(fig, ax, f"{cat['name']}: {MM} и {YY}, %", ax.get_lines() + ax2.get_lines())
            ytd_figure(cat["ytd"], f"{cat['name']}: накопленная с начала года инфляция, %",
                       cat["current_year"], str(int(cat["current_year"]) - 1), None)

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
