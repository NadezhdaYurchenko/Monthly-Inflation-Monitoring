# %%
"""
Месячный репортинг по инфляции.

Скрипт читает лист "Данные" из исходного Excel-файла, считает показатели в Python
и выгружает в новый Excel только готовые значения + редактируемые Excel-графики.
Опционально формирует PDF с графиками.

Зависимости:
    pip install pandas numpy openpyxl xlsxwriter matplotlib
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable
import difflib
import math
import re

import numpy as np
import pandas as pd


# =============================================================================
# 1. НАСТРОЙКИ — ЭТОТ БЛОК УДОБНО РЕДАКТИРОВАТЬ В JUPYTER
# =============================================================================

INPUT_FILE = Path("Месячная инфляция_Данные.xlsx")
DATA_SHEET = "Данные"

OUTPUT_XLSX = INPUT_FILE.with_name("Месячная инфляция_Отчет.xlsx")
OUTPUT_PDF = INPUT_FILE.with_name("Месячная инфляция_Графики.pdf")
EXPORT_PDF = True

# Быстрый график произвольной категории: достаточно поменять название здесь.
SELECTED_CATEGORY = "Мясопродукты"

# Проверка восстановления headline по группам 0 и 1 — только для последней даты.
VALIDATION_TOLERANCE_PP = 0.5

# Защита от расчета производных индексов на почти пустом месяце.
# В текущем файле обычное покрытие группы 0 составляет около 97.5–100% веса,
# поэтому 95% отсекает только явно неполные месяцы.
MIN_DETAILED_COVERAGE_WEIGHT = 95.0

# Показатель "без 20% наиболее волатильных компонентов".
VOLATILITY_WINDOW_MONTHS = 6
VOLATILITY_TRIM_SHARE = 0.20
# Минимум наблюдений внутри 6-месячного окна для оценки дисперсии.
# При строгом требовании шести наблюдений поставьте 6.
# Значение 4 делает расчет устойчивее к единичным пропускам.
VOLATILITY_MIN_OBS = 4

# Группы для рейтингов.
TOP_GROUPS = ["1", "11", "111", "0"]
TOP_N = 5

# Коды производных корзин.
FUEL_CODES = {"ТМ", "ТНМ"}
BASE_EXCLUSION_1 = {"РУ", "ПО", "С", "А", "ТМ", "ТНМ"}
BASE_EXCLUSION_2 = BASE_EXCLUSION_1 | {"ВТ", "ЗТ"}

# Названия ключевых колонок. Добавлены варианты написания, реально встречающиеся
# в текущем файле.
COLUMN_ALIASES = {
    "headline": ["Все товары и услуги"],
    "food": ["Продовольственные товары"],
    "nonfood": ["Непродовольственные товары"],
    "services": ["Услуги"],
    "sa_food": ["SA Продовольственые товары", "SA Продовольственные товары"],
    "sa_nonfood": [
        "SA Непродовольственная инфляция",
        "SA Непродовольственные товары",
    ],
    "sa_services": ["SA Услуги"],
}

MONTHS_RU = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}

# Цвета можно потом легко заменить в одном месте.
COLORS = {
    "blue": "#4472C4",
    "dark_blue": "#2F5597",
    "light_blue": "#D9EAF7",
    "orange": "#ED7D31",
    "green": "#70AD47",
    "red": "#C00000",
    "purple": "#7030A0",
    "grey": "#A5A5A5",
    "dark_grey": "#595959",
    "black": "#262626",
    "light_grey": "#E7E6E6",
}


# =============================================================================
# 2. СЛУЖЕБНЫЕ ФУНКЦИИ ЧТЕНИЯ И НОРМАЛИЗАЦИИ
# =============================================================================


def _is_na(value) -> bool:
    return value is None or (isinstance(value, float) and np.isnan(value)) or pd.isna(value)


def normalize_text(value) -> str:
    if _is_na(value):
        return ""
    text = str(value).replace("\xa0", " ").strip().lower()
    return re.sub(r"\s+", " ", text)


def normalize_group(value) -> str | None:
    if _is_na(value):
        return None
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    return str(value).strip().upper()


def normalize_code(value) -> str | None:
    if _is_na(value):
        return None
    text = str(value).strip().upper()
    return text if text else None


def parse_mixed_date(value) -> pd.Timestamp | pd.NaT:
    """Читает и Excel-datetime, и числовой serial date."""
    if _is_na(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return pd.Timestamp(value.year, value.month, 1)
    if isinstance(value, datetime):
        return pd.Timestamp(value.year, value.month, 1)
    if isinstance(value, date):
        return pd.Timestamp(value.year, value.month, 1)

    if isinstance(value, (int, float, np.integer, np.floating)):
        # Excel serial date. База 1899-12-30 учитывает особенности Excel.
        if float(value) > 20000:
            dt = pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
            return pd.Timestamp(dt.year, dt.month, 1)
        return pd.NaT

    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return pd.NaT
    return pd.Timestamp(parsed.year, parsed.month, 1)


@dataclass
class InflationInput:
    values: pd.DataFrame          # index = месяцы, columns = внутренние номера категорий
    meta: pd.DataFrame            # метаданные категорий
    weight_rows: dict[str, pd.Series]
    weights_by_date: pd.DataFrame
    raw: pd.DataFrame



def find_row_by_label(raw: pd.DataFrame, label: str) -> int:
    target = normalize_text(label)
    normalized = raw.iloc[:, 0].map(normalize_text)
    matches = normalized[normalized == target].index.tolist()
    if not matches:
        raise KeyError(f'Не найдена строка "{label}" на листе "{DATA_SHEET}".')
    return int(matches[0])



def weight_label_for_date(dt: pd.Timestamp) -> str:
    if dt.year == 2022:
        return "Вес 2022, до 01.04" if dt.month <= 3 else "Вес 2022, после 01.04"
    return f"Вес {dt.year}"



def load_input(path: Path) -> InflationInput:
    raw = pd.read_excel(path, sheet_name=DATA_SHEET, header=None, engine="openpyxl")
    if raw.shape[1] < 2:
        raise ValueError("На листе 'Данные' не найдены колонки с категориями.")

    n_categories = raw.shape[1] - 1
    cols = pd.RangeIndex(n_categories, name="category_id")

    # Заголовки категорий — первая строка, начиная с Excel-колонки B.
    names = raw.iloc[0, 1:].copy().reset_index(drop=True)
    names = names.map(lambda x: "" if _is_na(x) else str(x).strip())

    row_group = find_row_by_label(raw, "Группа")
    row_code = find_row_by_label(raw, "Код")
    row_fx = find_row_by_label(raw, "Вал.курс")

    # В текущем файле строка называется "Монетар. Инфляция".
    # Нормализация пробелов/регистра делает поиск устойчивым.
    row_monetary = find_row_by_label(raw, "Монетар. Инфляция")
    row_last_year = find_row_by_label(raw, "Последний год")
    row_rosstat = find_row_by_label(raw, "Код Росстат")

    meta = pd.DataFrame(index=cols)
    meta["name"] = names.values
    meta["name_norm"] = meta["name"].map(normalize_text)
    meta["rosstat_code"] = raw.iloc[row_rosstat, 1:].reset_index(drop=True).values
    meta["last_year"] = pd.to_numeric(
        raw.iloc[row_last_year, 1:].reset_index(drop=True), errors="coerce"
    )
    meta["group"] = raw.iloc[row_group, 1:].reset_index(drop=True).map(normalize_group).values
    meta["code"] = raw.iloc[row_code, 1:].reset_index(drop=True).map(normalize_code).values
    meta["fx"] = pd.to_numeric(
        raw.iloc[row_fx, 1:].reset_index(drop=True), errors="coerce"
    )
    meta["monetary"] = pd.to_numeric(
        raw.iloc[row_monetary, 1:].reset_index(drop=True), errors="coerce"
    )

    # Все строки весов читаются динамически. Добавление "Вес 2027" не требует правок кода.
    weight_rows: dict[str, pd.Series] = {}
    for idx, value in raw.iloc[:, 0].items():
        if _is_na(value):
            continue
        label = str(value).strip()
        if label.startswith("Вес "):
            weight_rows[label] = pd.Series(
                pd.to_numeric(raw.iloc[idx, 1:].reset_index(drop=True), errors="coerce").values,
                index=cols,
                dtype="float64",
            )

    # Строки данных определяются по датам в первой колонке.
    date_rows: list[int] = []
    dates: list[pd.Timestamp] = []
    for idx, value in raw.iloc[:, 0].items():
        dt = parse_mixed_date(value)
        if not pd.isna(dt):
            date_rows.append(int(idx))
            dates.append(dt)

    if not date_rows:
        raise ValueError("Не найдены строки с месячными датами.")

    values = raw.loc[date_rows, raw.columns[1:]].copy()
    values = values.apply(pd.to_numeric, errors="coerce")
    values.index = pd.DatetimeIndex(dates, name="Дата")
    values.columns = cols
    values = values.sort_index()

    if values.index.duplicated().any():
        duplicates = values.index[values.index.duplicated()].strftime("%Y-%m").tolist()
        raise ValueError(f"В данных есть дублирующиеся месяцы: {duplicates}")

    # Полная месячная сетка: это автоматически делает проверку 12 последовательных
    # наблюдений строгой — пропущенный месяц превращается в NaN.
    full_index = pd.date_range(values.index.min(), values.index.max(), freq="MS", name="Дата")
    values = values.reindex(full_index)

    # Матрица весов по каждой дате.
    weights_by_date = pd.DataFrame(index=values.index, columns=cols, dtype="float64")
    missing_weight_labels: set[str] = set()
    for dt in values.index:
        label = weight_label_for_date(dt)
        if label not in weight_rows:
            missing_weight_labels.add(label)
            continue
        weights_by_date.loc[dt] = weight_rows[label].values

    if missing_weight_labels:
        labels = ", ".join(sorted(missing_weight_labels))
        print(f"⚠️ Не найдены строки весов для: {labels}. Соответствующие расчеты будут NaN.")

    return InflationInput(
        values=values,
        meta=meta,
        weight_rows=weight_rows,
        weights_by_date=weights_by_date,
        raw=raw,
    )



def find_column(meta: pd.DataFrame, aliases: Iterable[str], *, required: bool = True) -> int | None:
    alias_norm = {normalize_text(x) for x in aliases}
    matches = meta.index[meta["name_norm"].isin(alias_norm)].tolist()

    if len(matches) == 1:
        return int(matches[0])
    if len(matches) > 1:
        names = meta.loc[matches, "name"].tolist()
        raise ValueError(f"Найдено несколько колонок для {list(aliases)}: {names}")

    if required:
        all_names = meta["name"].astype(str).tolist()
        suggestions = difflib.get_close_matches(list(aliases)[0], all_names, n=5, cutoff=0.5)
        raise KeyError(
            f"Не найдена колонка из вариантов {list(aliases)}. "
            f"Ближайшие названия: {suggestions}"
        )
    return None


# =============================================================================
# 3. РАСЧЕТНЫЕ ФУНКЦИИ
# =============================================================================


def monthly_index_to_yoy(index_series: pd.Series) -> pd.Series:
    """Г/г из последовательности месячных индексов вида 100.13."""
    factor = index_series.astype(float) / 100.0
    return (factor.rolling(12, min_periods=12).apply(np.prod, raw=True) - 1.0) * 100.0



def detailed_coverage_weight(inp: InflationInput) -> pd.Series:
    mask0 = inp.meta["group"].eq("0")
    coverage = pd.Series(index=inp.values.index, dtype="float64")
    for dt in inp.values.index:
        row = inp.values.loc[dt]
        weights = inp.weights_by_date.loc[dt]
        valid = mask0 & row.notna() & weights.notna()
        coverage.loc[dt] = float(weights[valid].sum())
    return coverage



def aggregate_group0_mm(
    inp: InflationInput,
    selector: pd.Series,
    coverage: pd.Series,
) -> pd.Series:
    """
    Агрегирует выбранную подкорзину группы 0.

    Возвращает м/м в процентах, а не индекс. Веса оставшейся подкорзины
    автоматически перенормируются на 100% делением на сумму ее исходных весов.
    """
    selector = selector.reindex(inp.meta.index).fillna(False).astype(bool)
    group0 = inp.meta["group"].eq("0")

    result = pd.Series(index=inp.values.index, dtype="float64")
    for dt in inp.values.index:
        if pd.isna(coverage.loc[dt]) or coverage.loc[dt] < MIN_DETAILED_COVERAGE_WEIGHT:
            result.loc[dt] = np.nan
            continue

        row = inp.values.loc[dt]
        weights = inp.weights_by_date.loc[dt]
        valid = group0 & selector & row.notna() & weights.notna()
        total_weight = float(weights[valid].sum())

        if total_weight <= 0:
            result.loc[dt] = np.nan
            continue

        # Индивидуальная м/м инфляция = индекс - 100.
        result.loc[dt] = float(((row[valid] - 100.0) * weights[valid]).sum() / total_weight)

    return result



def volatility_exclusion_mm(
    inp: InflationInput,
    coverage: pd.Series,
) -> tuple[pd.Series, pd.DataFrame]:
    """
    Исключает 20% наиболее волатильных категорий группы 0 по числу категорий.

    Для каждого месяца t дисперсия каждой категории считается по окну последних
    VOLATILITY_WINDOW_MONTHS месяцев, доступных к моменту t. То есть ряд строится
    рекурсивно/в псевдореальном времени, а не с использованием информации из будущего.

    Категории с недостаточным числом наблюдений для дисперсии не попадают в рейтинг
    волатильности и остаются в корзине.
    """
    group0 = inp.meta["group"].eq("0")
    result = pd.Series(index=inp.values.index, dtype="float64")
    diagnostics: list[dict] = []

    for pos, dt in enumerate(inp.values.index):
        if pd.isna(coverage.loc[dt]) or coverage.loc[dt] < MIN_DETAILED_COVERAGE_WEIGHT:
            result.loc[dt] = np.nan
            diagnostics.append({
                "Дата": dt,
                "Категорий с текущими данными": np.nan,
                "Категорий с оцененной дисперсией": np.nan,
                "Исключено категорий": np.nan,
                "Вес оставшейся корзины": np.nan,
            })
            continue

        row = inp.values.loc[dt]
        weights = inp.weights_by_date.loc[dt]
        current_valid = group0 & row.notna() & weights.notna()
        current_cols = inp.meta.index[current_valid]

        start = max(0, pos - VOLATILITY_WINDOW_MONTHS + 1)
        hist = inp.values.iloc[start:pos + 1, current_cols] - 100.0
        obs_count = hist.count(axis=0)
        variances = hist.var(axis=0, ddof=1, skipna=True)
        eligible = variances[(obs_count >= VOLATILITY_MIN_OBS) & variances.notna()]

        n_trim = int(math.ceil(VOLATILITY_TRIM_SHARE * len(eligible))) if len(eligible) else 0
        trimmed_cols = set(eligible.sort_values(ascending=False).head(n_trim).index.tolist())
        kept_cols = [c for c in current_cols if c not in trimmed_cols]

        kept_weights = weights.loc[kept_cols]
        total_weight = float(kept_weights.sum())
        if total_weight <= 0:
            result.loc[dt] = np.nan
        else:
            result.loc[dt] = float(
                ((row.loc[kept_cols] - 100.0) * kept_weights).sum() / total_weight
            )

        diagnostics.append({
            "Дата": dt,
            "Категорий с текущими данными": int(len(current_cols)),
            "Категорий с оцененной дисперсией": int(len(eligible)),
            "Исключено категорий": int(n_trim),
            "Вес оставшейся корзины": total_weight,
        })

    diag_df = pd.DataFrame(diagnostics).set_index("Дата")
    return result, diag_df



def latest_category_yoy(inp: InflationInput, latest: pd.Timestamp) -> pd.Series:
    """Г/г по каждой категории. Нужны 12 последовательных месячных наблюдений."""
    all_yoy = (inp.values / 100.0).rolling(12, min_periods=12).apply(np.prod, raw=True)
    all_yoy = (all_yoy - 1.0) * 100.0
    return all_yoy.loc[latest]



def top_positive_negative(series: pd.Series, n: int = TOP_N) -> tuple[pd.Series, pd.Series]:
    clean = series.replace([np.inf, -np.inf], np.nan).dropna()
    positive = clean[clean > 0].sort_values(ascending=False).head(n)
    negative = clean[clean < 0].sort_values(ascending=True).head(n)
    return positive, negative



def cumulative_ytd(index_series: pd.Series, year: int, through_month: int) -> pd.Series:
    idx = index_series[index_series.index.year == year].copy()
    idx = idx[idx.index.month <= through_month]
    expected = pd.date_range(f"{year}-01-01", periods=through_month, freq="MS")
    idx = idx.reindex(expected)
    return ((idx / 100.0).cumprod() - 1.0) * 100.0


# =============================================================================
# 4. ПРОВЕРКИ — ПЕЧАТАЮТСЯ ТОЛЬКО В JUPYTER / КОНСОЛЬ
# =============================================================================


def print_last_observation_checks(
    inp: InflationInput,
    cols: dict[str, int],
    coverage: pd.Series,
) -> pd.Timestamp:
    headline = inp.values[cols["headline"]].dropna()
    if headline.empty:
        raise ValueError("В колонке headline нет данных.")
    latest = headline.index.max()

    weights = inp.weights_by_date.loc[latest]
    row = inp.values.loc[latest]
    headline_index = float(row[cols["headline"]])

    print("=" * 78)
    print(f"Последнее наблюдение: {latest:%d.%m.%Y}")
    print(f"Используемая строка весов: {weight_label_for_date(latest)}")
    print(f"Покрытие группы 0 доступными данными и весами: {coverage.loc[latest]:.3f}%")
    print("-" * 78)

    for group in ["0", "1"]:
        mask = inp.meta["group"].eq(group)
        valid = mask & row.notna() & weights.notna()
        reconstructed = 100.0 + float(((row[valid] - 100.0) * weights[valid]).sum() / 100.0)
        error_pp = reconstructed - headline_index
        status = "✅ OK" if abs(error_pp) < VALIDATION_TOLERANCE_PP else "⚠️ WARNING"
        print(
            f"{status} | Группа {group}: headline={headline_index:.4f}, "
            f"расчет={reconstructed:.4f}, ошибка={error_pp:+.4f} п.п. "
            f"(допуск ±{VALIDATION_TOLERANCE_PP:.2f} п.п.)"
        )

    # Короткий контроль дыр в детальных данных после начала детальной истории.
    positive_coverage = coverage[coverage > 0]
    if not positive_coverage.empty:
        first_detailed = positive_coverage.index.min()
        bad = coverage.loc[first_detailed:]
        bad = bad[bad < MIN_DETAILED_COVERAGE_WEIGHT]
        if not bad.empty:
            pairs = ", ".join(f"{dt:%m.%Y} ({value:.1f}% веса)" for dt, value in bad.items())
            print(
                f"⚠️ Неполные детальные месяцы ниже порога {MIN_DETAILED_COVERAGE_WEIGHT:.0f}%: "
                f"{pairs}. Производные индексы за эти месяцы будут NaN."
            )

    # Контроль разметки монетарной/немонетарной корзины на последней дате.
    active0 = inp.meta["group"].eq("0") & row.notna() & weights.notna()
    unclassified = active0 & inp.meta["monetary"].isna()
    unclassified_weight = float(weights[unclassified].sum())
    if unclassified_weight > 1e-9:
        print(
            f"⚠️ В группе 0 на последней дате есть неразмеченный по монетарности вес: "
            f"{unclassified_weight:.3f}%"
        )

    print("=" * 78)
    return latest


# =============================================================================
# 5. РАСЧЕТ ВСЕХ ПОКАЗАТЕЛЕЙ
# =============================================================================


@dataclass
class ReportData:
    latest: pd.Timestamp
    main: pd.DataFrame
    sa_contributions: pd.DataFrame
    derived: pd.DataFrame
    base: pd.DataFrame
    ytd: pd.DataFrame
    selected_category: pd.DataFrame
    selected_category_name: str
    top_data: dict
    volatility_diagnostics: pd.DataFrame



def calculate_report(inp: InflationInput) -> ReportData:
    cols = {key: find_column(inp.meta, aliases) for key, aliases in COLUMN_ALIASES.items()}
    coverage = detailed_coverage_weight(inp)
    latest = print_last_observation_checks(inp, cols, coverage)

    # -------------------------------------------------------------------------
    # Основной headline
    # -------------------------------------------------------------------------
    headline_index = inp.values[cols["headline"]]
    headline_mm = headline_index - 100.0
    headline_yoy = monthly_index_to_yoy(headline_index)

    # SA headline строим в Python из трех уже сезонно сглаженных агрегатов.
    # Никакой сезонной корректировки внутри Python здесь не происходит.
    sa_food_mm = inp.values[cols["sa_food"]] - 100.0
    sa_nonfood_mm = inp.values[cols["sa_nonfood"]] - 100.0
    sa_services_mm = inp.values[cols["sa_services"]] - 100.0

    w_food = inp.weights_by_date[cols["food"]]
    w_nonfood = inp.weights_by_date[cols["nonfood"]]
    w_services = inp.weights_by_date[cols["services"]]

    contrib_food = sa_food_mm * w_food / 100.0
    contrib_nonfood = sa_nonfood_mm * w_nonfood / 100.0
    contrib_services = sa_services_mm * w_services / 100.0

    headline_sa_mm = contrib_food + contrib_nonfood + contrib_services
    headline_sa_index = 100.0 + headline_sa_mm
    headline_saar = ((headline_sa_index / 100.0) ** 12 - 1.0) * 100.0
    headline_3mma_saar = headline_saar.rolling(3, min_periods=3).mean()

    main = pd.DataFrame({
        "Дата": inp.values.index,
        "Headline SA м/м, %": headline_sa_mm.values,
        "Headline SAAR, %": headline_saar.values,
        "3mma SAAR, %": headline_3mma_saar.values,
        "Headline г/г, %": headline_yoy.values,
    })

    sa_contributions = pd.DataFrame({
        "Дата": inp.values.index,
        "Headline SA м/м, %": headline_sa_mm.values,
        "Вклад продовольствия, п.п.": contrib_food.values,
        "Вклад непродовольственных товаров, п.п.": contrib_nonfood.values,
        "Вклад услуг, п.п.": contrib_services.values,
    })

    # -------------------------------------------------------------------------
    # Производные корзины группы 0
    # -------------------------------------------------------------------------
    code = inp.meta["code"]
    monetary_flag = inp.meta["monetary"]

    no_fuel_mm = aggregate_group0_mm(inp, ~code.isin(FUEL_CODES), coverage)
    monetary_mm = aggregate_group0_mm(inp, monetary_flag.eq(1), coverage)
    nonmonetary_mm = aggregate_group0_mm(inp, monetary_flag.eq(0), coverage)

    no_fuel_yoy = monthly_index_to_yoy(100.0 + no_fuel_mm)
    monetary_yoy = monthly_index_to_yoy(100.0 + monetary_mm)
    nonmonetary_yoy = monthly_index_to_yoy(100.0 + nonmonetary_mm)

    derived = pd.DataFrame({
        "Дата": inp.values.index,
        "Без топлива м/м, %": no_fuel_mm.values,
        "Без топлива г/г, %": no_fuel_yoy.values,
        "Монетарная м/м, %": monetary_mm.values,
        "Монетарная г/г, %": monetary_yoy.values,
        "Немонетарная м/м, %": nonmonetary_mm.values,
        "Немонетарная г/г, %": nonmonetary_yoy.values,
    })

    # -------------------------------------------------------------------------
    # Базовая / устойчивая инфляция: три статистических варианта
    # -------------------------------------------------------------------------
    volatile_mm, volatility_diagnostics = volatility_exclusion_mm(inp, coverage)
    base_excl_1_mm = aggregate_group0_mm(inp, ~code.isin(BASE_EXCLUSION_1), coverage)
    base_excl_2_mm = aggregate_group0_mm(inp, ~code.isin(BASE_EXCLUSION_2), coverage)

    volatile_yoy = monthly_index_to_yoy(100.0 + volatile_mm)
    base_excl_1_yoy = monthly_index_to_yoy(100.0 + base_excl_1_mm)
    base_excl_2_yoy = monthly_index_to_yoy(100.0 + base_excl_2_mm)

    base = pd.DataFrame({
        "Дата": inp.values.index,
        "Без 20% наиболее волатильных, м/м, %": volatile_mm.values,
        "Без 20% наиболее волатильных, г/г, %": volatile_yoy.values,
        "Исключение [1], м/м, %": base_excl_1_mm.values,
        "Исключение [1], г/г, %": base_excl_1_yoy.values,
        "Исключение [2], м/м, %": base_excl_2_mm.values,
        "Исключение [2], г/г, %": base_excl_2_yoy.values,
    })

    # -------------------------------------------------------------------------
    # Накопленная инфляция: текущий год vs тот же период прошлого года
    # -------------------------------------------------------------------------
    current_year = latest.year
    previous_year = current_year - 1
    through_month = latest.month

    ytd_current = cumulative_ytd(headline_index, current_year, through_month)
    ytd_previous = cumulative_ytd(headline_index, previous_year, through_month)

    ytd = pd.DataFrame({
        "Месяц": [MONTHS_RU[m] for m in range(1, through_month + 1)],
        f"{current_year}, накопленно с начала года, %": ytd_current.values,
        f"{previous_year}, накопленно с начала года, %": ytd_previous.values,
    })

    # -------------------------------------------------------------------------
    # Произвольная категория
    # -------------------------------------------------------------------------
    selected_col = find_column(inp.meta, [SELECTED_CATEGORY], required=False)
    if selected_col is None:
        suggestions = difflib.get_close_matches(
            SELECTED_CATEGORY, inp.meta["name"].astype(str).tolist(), n=7, cutoff=0.45
        )
        print(
            f"⚠️ Категория '{SELECTED_CATEGORY}' не найдена. "
            f"Ближайшие варианты: {suggestions}. Лист 'Категория' будет пустым."
        )
        selected_name = SELECTED_CATEGORY
        selected_category = pd.DataFrame(columns=["Дата", "м/м, %", "г/г, %"])
    else:
        selected_name = str(inp.meta.loc[selected_col, "name"])
        selected_index = inp.values[selected_col]
        selected_category = pd.DataFrame({
            "Дата": inp.values.index,
            "м/м, %": (selected_index - 100.0).values,
            "г/г, %": monthly_index_to_yoy(selected_index).values,
        })

    # -------------------------------------------------------------------------
    # Топ-5 для последней даты
    # -------------------------------------------------------------------------
    current_row = inp.values.loc[latest]
    current_weights = inp.weights_by_date.loc[latest]
    yoy_by_category = latest_category_yoy(inp, latest)
    mm_by_category = current_row - 100.0
    contrib_by_category = (current_row - 100.0) * current_weights / 100.0

    top_data: dict = {}
    for group in TOP_GROUPS:
        mask = inp.meta["group"].eq(group)
        group_cols = inp.meta.index[mask]

        metrics = {
            "м/м, %": mm_by_category.loc[group_cols],
            "г/г, %": yoy_by_category.loc[group_cols],
            "Вклад в headline, п.п.": contrib_by_category.loc[group_cols],
        }
        top_data[group] = {}
        for metric_name, metric_series in metrics.items():
            metric_series = metric_series.copy()
            metric_series.index = inp.meta.loc[metric_series.index, "name"].values
            top_data[group][metric_name] = top_positive_negative(metric_series, TOP_N)

    # Текущая диагностика варианта по волатильности — только в Jupyter.
    if latest in volatility_diagnostics.index:
        d = volatility_diagnostics.loc[latest]
        if d.notna().any():
            print(
                "Показатель без 20% наиболее волатильных на последней дате: "
                f"исключено {int(d['Исключено категорий'])} категорий; "
                f"вес оставшейся корзины до перенормировки = "
                f"{d['Вес оставшейся корзины']:.3f}%."
            )

    return ReportData(
        latest=latest,
        main=main,
        sa_contributions=sa_contributions,
        derived=derived,
        base=base,
        ytd=ytd,
        selected_category=selected_category,
        selected_category_name=selected_name,
        top_data=top_data,
        volatility_diagnostics=volatility_diagnostics,
    )


# =============================================================================
# 6. EXCEL: ТОЛЬКО ГОТОВЫЕ ЗНАЧЕНИЯ + РЕДАКТИРУЕМЫЕ ГРАФИКИ
# =============================================================================


def _chart_base_format(chart, title: str, y_title: str = "%") -> None:
    chart.set_title({"name": title, "name_font": {"name": "Arial", "size": 14, "bold": True}})
    chart.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 9}})
    chart.set_x_axis({
        "date_axis": True,
        "num_format": "mmm-yy",
        "label_position": "low",
        "num_font": {"name": "Arial", "size": 8},
        "line": {"color": "#BFBFBF"},
        "major_gridlines": {"visible": False},
    })
    chart.set_y_axis({
        "name": y_title,
        "name_font": {"name": "Arial", "size": 9},
        "num_font": {"name": "Arial", "size": 8},
        "major_gridlines": {"visible": True, "line": {"color": "#E7E6E6"}},
        "line": {"color": "#BFBFBF"},
    })
    chart.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
    chart.set_plotarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
    chart.set_size({"width": 900, "height": 430})



def _format_dataframe_sheet(writer, sheet_name: str, df: pd.DataFrame) -> None:
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    font = workbook.add_format({"font_name": "Arial"})
    header = workbook.add_format({
        "font_name": "Arial", "bold": True, "font_color": "#FFFFFF",
        "bg_color": COLORS["dark_blue"], "border": 0,
        "align": "center", "valign": "vcenter", "text_wrap": True,
    })
    date_fmt = workbook.add_format({"font_name": "Arial", "num_format": "dd.mm.yyyy"})
    num_fmt = workbook.add_format({"font_name": "Arial", "num_format": "0.00"})

    worksheet.hide_gridlines(2)
    worksheet.freeze_panes(1, 1)
    worksheet.set_row(0, 32, header)

    # Перезаписываем формат шапки без изменения значений.
    for c, value in enumerate(df.columns):
        worksheet.write(0, c, value, header)

    for c, col in enumerate(df.columns):
        if normalize_text(col) == "дата":
            worksheet.set_column(c, c, 13, date_fmt)
        elif "месяц" in normalize_text(col):
            worksheet.set_column(c, c, 14, font)
        else:
            width = min(max(16, len(str(col)) * 0.9), 34)
            worksheet.set_column(c, c, width, num_fmt)



def _add_main_chart(writer, df: pd.DataFrame, sheet_name: str) -> None:
    wb = writer.book
    ws = writer.sheets[sheet_name]
    n = len(df)
    if n == 0:
        return

    area = wb.add_chart({"type": "area"})
    area.add_series({
        "name": "Headline г/г",
        "categories": [sheet_name, 1, 0, n, 0],
        "values": [sheet_name, 1, 4, n, 4],
        "fill": {"color": COLORS["light_blue"], "transparency": 25},
        "line": {"color": COLORS["light_blue"], "transparency": 100},
    })

    line = wb.add_chart({"type": "line"})
    line.add_series({
        "name": "Headline SAAR",
        "categories": [sheet_name, 1, 0, n, 0],
        "values": [sheet_name, 1, 2, n, 2],
        "line": {"color": COLORS["dark_blue"], "width": 2.25},
    })
    line.add_series({
        "name": "3mma SAAR",
        "categories": [sheet_name, 1, 0, n, 0],
        "values": [sheet_name, 1, 3, n, 3],
        "line": {"none": True},
        "marker": {
            "type": "circle", "size": 5,
            "border": {"color": COLORS["orange"]},
            "fill": {"color": COLORS["orange"]},
        },
    })
    area.combine(line)
    _chart_base_format(area, "Headline: SAAR, 3mma SAAR и инфляция г/г")
    ws.insert_chart("G2", area)



def _add_contributions_chart(writer, df: pd.DataFrame, sheet_name: str) -> None:
    wb = writer.book
    ws = writer.sheets[sheet_name]
    n = len(df)
    if n == 0:
        return

    column = wb.add_chart({"type": "column", "subtype": "stacked"})
    series_specs = [
        (2, "Продовольствие", COLORS["orange"]),
        (3, "Непродовольственные товары", COLORS["green"]),
        (4, "Услуги", COLORS["blue"]),
    ]
    for col_idx, name, color in series_specs:
        column.add_series({
            "name": name,
            "categories": [sheet_name, 1, 0, n, 0],
            "values": [sheet_name, 1, col_idx, n, col_idx],
            "fill": {"color": color},
            "border": {"none": True},
        })

    line = wb.add_chart({"type": "line"})
    line.add_series({
        "name": "Headline SA м/м",
        "categories": [sheet_name, 1, 0, n, 0],
        "values": [sheet_name, 1, 1, n, 1],
        "line": {"color": COLORS["black"], "width": 2.0},
    })
    column.combine(line)
    _chart_base_format(column, "Вклады продовольствия, непродовольственных товаров и услуг в SA headline")
    ws.insert_chart("G2", column)



def _add_line_chart(
    writer,
    sheet_name: str,
    df: pd.DataFrame,
    columns: list[tuple[int, str, str]],
    title: str,
    position: str,
    y_title: str = "%",
    secondary_columns: set[int] | None = None,
) -> None:
    wb = writer.book
    ws = writer.sheets[sheet_name]
    n = len(df)
    if n == 0:
        return
    secondary_columns = secondary_columns or set()

    chart = wb.add_chart({"type": "line"})
    for col_idx, name, color in columns:
        chart.add_series({
            "name": name,
            "categories": [sheet_name, 1, 0, n, 0],
            "values": [sheet_name, 1, col_idx, n, col_idx],
            "line": {"color": color, "width": 2.0},
            "y2_axis": col_idx in secondary_columns,
        })
    _chart_base_format(chart, title, y_title)
    if secondary_columns:
        chart.set_y2_axis({
            "name": "%",
            "name_font": {"name": "Arial", "size": 9},
            "num_font": {"name": "Arial", "size": 8},
            "major_gridlines": {"visible": False},
        })
    ws.insert_chart(position, chart)



def _write_top_sheet(writer, report: ReportData) -> None:
    wb = writer.book
    sheet_name = "Топ-5"
    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws
    ws.hide_gridlines(2)

    title_fmt = wb.add_format({
        "font_name": "Arial", "bold": True, "font_size": 14,
        "font_color": COLORS["dark_blue"],
    })
    section_fmt = wb.add_format({
        "font_name": "Arial", "bold": True, "font_size": 11,
        "font_color": "#FFFFFF", "bg_color": COLORS["dark_blue"],
        "align": "left",
    })
    header_fmt = wb.add_format({
        "font_name": "Arial", "bold": True, "bg_color": COLORS["light_grey"],
        "align": "center", "text_wrap": True,
    })
    text_fmt = wb.add_format({"font_name": "Arial"})
    num_fmt = wb.add_format({"font_name": "Arial", "num_format": "0.00"})

    ws.set_column("A:A", 6, text_fmt)
    ws.set_column("B:B", 44, text_fmt)
    ws.set_column("C:C", 13, num_fmt)
    ws.set_column("E:E", 6, text_fmt)
    ws.set_column("F:F", 44, text_fmt)
    ws.set_column("G:G", 13, num_fmt)

    ws.write("A1", f"Топ-5 категорий на {report.latest:%d.%m.%Y}", title_fmt)

    row = 2
    metric_order = ["м/м, %", "г/г, %", "Вклад в headline, п.п."]
    for group in TOP_GROUPS:
        ws.merge_range(row, 0, row, 6, f"Группа {group}", section_fmt)
        row += 1
        for metric in metric_order:
            pos, neg = report.top_data[group][metric]
            ws.merge_range(row, 0, row, 6, metric, header_fmt)
            row += 1
            ws.write_row(row, 0, ["Место", "Сильнее всего выросли", "Значение"], header_fmt)
            ws.write_row(row, 4, ["Место", "Сильнее всего снизились", "Значение"], header_fmt)
            row += 1

            for rank in range(TOP_N):
                if rank < len(pos):
                    ws.write(row + rank, 0, rank + 1, text_fmt)
                    ws.write(row + rank, 1, str(pos.index[rank]), text_fmt)
                    ws.write(row + rank, 2, float(pos.iloc[rank]), num_fmt)
                if rank < len(neg):
                    ws.write(row + rank, 4, rank + 1, text_fmt)
                    ws.write(row + rank, 5, str(neg.index[rank]), text_fmt)
                    ws.write(row + rank, 6, float(neg.iloc[rank]), num_fmt)
            row += TOP_N + 2



def export_excel(report: ReportData, path: Path) -> None:
    with pd.ExcelWriter(
        path,
        engine="xlsxwriter",
        datetime_format="dd.mm.yyyy",
        date_format="dd.mm.yyyy",
    ) as writer:
        # Значения — без формул.
        report.main.to_excel(writer, sheet_name="Основное", index=False)
        report.sa_contributions.to_excel(writer, sheet_name="Вклады SA", index=False)
        report.derived.to_excel(writer, sheet_name="Производные", index=False)
        report.base.to_excel(writer, sheet_name="Базовая", index=False)
        report.ytd.to_excel(writer, sheet_name="Накопленная", index=False)
        report.selected_category.to_excel(writer, sheet_name="Категория", index=False)

        for sheet_name, df in [
            ("Основное", report.main),
            ("Вклады SA", report.sa_contributions),
            ("Производные", report.derived),
            ("Базовая", report.base),
            ("Накопленная", report.ytd),
            ("Категория", report.selected_category),
        ]:
            _format_dataframe_sheet(writer, sheet_name, df)

        _add_main_chart(writer, report.main, "Основное")
        _add_contributions_chart(writer, report.sa_contributions, "Вклады SA")

        # Производные: три графика.
        _add_line_chart(
            writer, "Производные", report.derived,
            [(1, "Без топлива м/м", COLORS["orange"]),
             (2, "Без топлива г/г", COLORS["dark_blue"])],
            "Инфляция без топлива: м/м и г/г",
            "I2",
            secondary_columns={1},
        )
        _add_line_chart(
            writer, "Производные", report.derived,
            [(3, "Монетарная м/м", COLORS["red"]),
             (5, "Немонетарная м/м", COLORS["green"])],
            "Монетарная и немонетарная инфляция: м/м",
            "I24",
        )
        _add_line_chart(
            writer, "Производные", report.derived,
            [(4, "Монетарная г/г", COLORS["red"]),
             (6, "Немонетарная г/г", COLORS["green"])],
            "Монетарная и немонетарная инфляция: г/г",
            "I46",
        )

        # Базовая: отдельно м/м и г/г.
        _add_line_chart(
            writer, "Базовая", report.base,
            [(1, "Без 20% наиболее волатильных", COLORS["purple"]),
             (3, "Исключение [1]", COLORS["orange"]),
             (5, "Исключение [2]", COLORS["dark_blue"])],
            "Базовая инфляция: м/м",
            "I2",
        )
        _add_line_chart(
            writer, "Базовая", report.base,
            [(2, "Без 20% наиболее волатильных", COLORS["purple"]),
             (4, "Исключение [1]", COLORS["orange"]),
             (6, "Исключение [2]", COLORS["dark_blue"])],
            "Базовая инфляция: г/г",
            "I24",
        )

        # Накопленная инфляция.
        wb = writer.book
        ws_ytd = writer.sheets["Накопленная"]
        if len(report.ytd):
            chart = wb.add_chart({"type": "line"})
            for col_idx, color in [(1, COLORS["dark_blue"]), (2, COLORS["grey"])]:
                chart.add_series({
                    "name": ["Накопленная", 0, col_idx],
                    "categories": ["Накопленная", 1, 0, len(report.ytd), 0],
                    "values": ["Накопленная", 1, col_idx, len(report.ytd), col_idx],
                    "line": {"color": color, "width": 2.25},
                    "marker": {"type": "circle", "size": 4},
                })
            chart.set_title({
                "name": "Накопленная инфляция с начала года",
                "name_font": {"name": "Arial", "size": 14, "bold": True},
            })
            chart.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 9}})
            chart.set_x_axis({"num_font": {"name": "Arial", "size": 9}})
            chart.set_y_axis({
                "name": "%", "num_font": {"name": "Arial", "size": 8},
                "major_gridlines": {"visible": True, "line": {"color": "#E7E6E6"}},
            })
            chart.set_chartarea({"border": {"none": True}})
            chart.set_plotarea({"border": {"none": True}})
            chart.set_size({"width": 800, "height": 420})
            ws_ytd.insert_chart("E2", chart)

        # Выбранная категория.
        if len(report.selected_category):
            _add_line_chart(
                writer, "Категория", report.selected_category,
                [(1, f"{report.selected_category_name}: м/м", COLORS["orange"]),
                 (2, f"{report.selected_category_name}: г/г", COLORS["dark_blue"])],
                f"{report.selected_category_name}: м/м и г/г",
                "E2",
                secondary_columns={1},
            )

        _write_top_sheet(writer, report)

    print(f"✅ Excel сохранен: {path.resolve()}")


# =============================================================================
# 7. PDF: ОПЦИОНАЛЬНО, ПО ОДНОМУ ГРАФИКУ НА СТРАНИЦУ
# =============================================================================


def export_pdf(report: ReportData, path: Path) -> None:
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.backends.backend_pdf import PdfPages

    plt.rcParams["font.family"] = "Arial"
    plt.rcParams["axes.spines.top"] = False
    plt.rcParams["axes.spines.right"] = False

    def finish(ax, title: str, ylabel: str = "%"):
        ax.set_title(title, fontsize=15, fontweight="bold")
        ax.set_ylabel(ylabel)
        ax.grid(axis="y", alpha=0.25)
        ax.legend(frameon=False, loc="best")
        if np.issubdtype(np.asarray(ax.get_xticks()).dtype, np.number):
            pass
        fig = ax.figure
        fig.tight_layout()
        return fig

    with PdfPages(path) as pdf:
        # 1. Headline.
        d = report.main.dropna(subset=["Headline SAAR, %", "Headline г/г, %"], how="all")
        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        ax.fill_between(d["Дата"], d["Headline г/г, %"], alpha=0.22,
                        color=COLORS["light_blue"], label="Headline г/г")
        ax.plot(d["Дата"], d["Headline SAAR, %"], color=COLORS["dark_blue"],
                linewidth=2.2, label="Headline SAAR")
        ax.scatter(d["Дата"], d["3mma SAAR, %"], color=COLORS["orange"],
                   s=20, label="3mma SAAR", zorder=3)
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        pdf.savefig(finish(ax, "Headline: SAAR, 3mma SAAR и инфляция г/г"))
        plt.close(fig)

        # 2. Вклады SA — положительные и отрицательные стеки отдельно.
        d = report.sa_contributions.dropna(subset=[
            "Вклад продовольствия, п.п.",
            "Вклад непродовольственных товаров, п.п.",
            "Вклад услуг, п.п.",
        ], how="all")
        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        series = [
            ("Вклад продовольствия, п.п.", COLORS["orange"], "Продовольствие"),
            ("Вклад непродовольственных товаров, п.п.", COLORS["green"], "Непродовольственные товары"),
            ("Вклад услуг, п.п.", COLORS["blue"], "Услуги"),
        ]
        pos_bottom = np.zeros(len(d))
        neg_bottom = np.zeros(len(d))
        for col, color, label in series:
            vals = d[col].fillna(0).to_numpy(dtype=float)
            bottom = np.where(vals >= 0, pos_bottom, neg_bottom)
            ax.bar(d["Дата"], vals, bottom=bottom, width=20, color=color, label=label)
            pos_bottom += np.where(vals > 0, vals, 0)
            neg_bottom += np.where(vals < 0, vals, 0)
        ax.plot(d["Дата"], d["Headline SA м/м, %"], color=COLORS["black"],
                linewidth=1.8, label="Headline SA м/м")
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        pdf.savefig(finish(ax, "Вклады в SA headline", "п.п."))
        plt.close(fig)

        # 3. Без топлива.
        d = report.derived
        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        ax.plot(d["Дата"], d["Без топлива г/г, %"], color=COLORS["dark_blue"],
                linewidth=2.2, label="Без топлива г/г")
        ax2 = ax.twinx()
        ax2.plot(d["Дата"], d["Без топлива м/м, %"], color=COLORS["orange"],
                 linewidth=1.5, label="Без топлива м/м")
        ax2.set_ylabel("м/м, %")
        lines = ax.get_lines() + ax2.get_lines()
        ax.legend(lines, [x.get_label() for x in lines], frameon=False)
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        ax.set_title("Инфляция без топлива: м/м и г/г", fontsize=15, fontweight="bold")
        ax.set_ylabel("г/г, %")
        ax.grid(axis="y", alpha=0.25)
        fig.tight_layout()
        pdf.savefig(fig)
        plt.close(fig)

        # 4. Монетарная / немонетарная м/м.
        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        ax.plot(d["Дата"], d["Монетарная м/м, %"], color=COLORS["red"], label="Монетарная")
        ax.plot(d["Дата"], d["Немонетарная м/м, %"], color=COLORS["green"], label="Немонетарная")
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        pdf.savefig(finish(ax, "Монетарная и немонетарная инфляция: м/м"))
        plt.close(fig)

        # 5. Монетарная / немонетарная г/г.
        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        ax.plot(d["Дата"], d["Монетарная г/г, %"], color=COLORS["red"], label="Монетарная")
        ax.plot(d["Дата"], d["Немонетарная г/г, %"], color=COLORS["green"], label="Немонетарная")
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        pdf.savefig(finish(ax, "Монетарная и немонетарная инфляция: г/г"))
        plt.close(fig)

        # 6–7. Базовая м/м и г/г.
        b = report.base
        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        ax.plot(b["Дата"], b["Без 20% наиболее волатильных, м/м, %"], color=COLORS["purple"],
                label="Без 20% наиболее волатильных")
        ax.plot(b["Дата"], b["Исключение [1], м/м, %"], color=COLORS["orange"], label="Исключение [1]")
        ax.plot(b["Дата"], b["Исключение [2], м/м, %"], color=COLORS["dark_blue"], label="Исключение [2]")
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        pdf.savefig(finish(ax, "Базовая инфляция: м/м"))
        plt.close(fig)

        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        ax.plot(b["Дата"], b["Без 20% наиболее волатильных, г/г, %"], color=COLORS["purple"],
                label="Без 20% наиболее волатильных")
        ax.plot(b["Дата"], b["Исключение [1], г/г, %"], color=COLORS["orange"], label="Исключение [1]")
        ax.plot(b["Дата"], b["Исключение [2], г/г, %"], color=COLORS["dark_blue"], label="Исключение [2]")
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        pdf.savefig(finish(ax, "Базовая инфляция: г/г"))
        plt.close(fig)

        # 8. Накопленная.
        fig, ax = plt.subplots(figsize=(11.7, 7.0))
        ax.plot(report.ytd["Месяц"], report.ytd.iloc[:, 1], marker="o",
                color=COLORS["dark_blue"], label=report.ytd.columns[1])
        ax.plot(report.ytd["Месяц"], report.ytd.iloc[:, 2], marker="o",
                color=COLORS["grey"], label=report.ytd.columns[2])
        ax.tick_params(axis="x", rotation=35)
        pdf.savefig(finish(ax, "Накопленная инфляция с начала года"))
        plt.close(fig)

        # 9. Выбранная категория.
        if len(report.selected_category):
            c = report.selected_category
            fig, ax = plt.subplots(figsize=(11.7, 7.0))
            ax.plot(c["Дата"], c["г/г, %"], color=COLORS["dark_blue"],
                    label=f"{report.selected_category_name}: г/г")
            ax2 = ax.twinx()
            ax2.plot(c["Дата"], c["м/м, %"], color=COLORS["orange"],
                     label=f"{report.selected_category_name}: м/м")
            ax2.set_ylabel("м/м, %")
            lines = ax.get_lines() + ax2.get_lines()
            ax.legend(lines, [x.get_label() for x in lines], frameon=False)
            ax.xaxis.set_major_locator(mdates.YearLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
            ax.set_title(report.selected_category_name, fontsize=15, fontweight="bold")
            ax.set_ylabel("г/г, %")
            ax.grid(axis="y", alpha=0.25)
            fig.tight_layout()
            pdf.savefig(fig)
            plt.close(fig)

    print(f"✅ PDF сохранен: {path.resolve()}")


# =============================================================================
# 8. ЗАПУСК
# =============================================================================


def main() -> ReportData:
    inp = load_input(INPUT_FILE)
    report = calculate_report(inp)
    export_excel(report, OUTPUT_XLSX)
    if EXPORT_PDF:
        export_pdf(report, OUTPUT_PDF)
    return report


# В Jupyter можно либо выполнить весь файл, либо закомментировать этот блок
# и вызвать вручную: report = main()
if __name__ == "__main__":
    report = main()
