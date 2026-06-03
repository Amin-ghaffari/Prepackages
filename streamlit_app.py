# streamlit_app.py
import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
from openpyxl import Workbook
import math, re, statistics, warnings
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from scipy import stats

# =========================
# تنظیمات اصلی
# =========================
INPUT_DIR = r"."
OUTPUT_FILE = "گزارش_نهایی_پایدار.xlsx"
DEFAULT_NOMINAL = None
RECURSIVE = True
MIN_ROWS_FOR_ANALYSIS = 20

PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

COLUMN_KEYWORDS = {
    "gross": ["وزنناخالص", "gross"],
    "package": ["وزنبسته", "بستهها", "tare", "package"],
    "net_weight": ["وزنخالص", "netweight"],
    "net_volume": ["حجمخالص", "netvolume"],
    "corrected_weight": ["وزنخالصتصحیح", "correctedweight", "وزن تصحیح"],
    "corrected_volume": ["حجمخالصتصحیح", "correctedvolume", "حجم تصحیح"],
    "t1_error": ["خطایt1", "error1", "t1"],
    "t2_error": ["خطایt2", "error2", "t2"],
}

SUMMARY_PATTERNS = [
    ("nominal", ["وزننامی", "حجمنامی", "مقدارنامی"]),
    ("tolerance", ["رواداریمجاز"]),
    ("density", ["چگالی"]),
    ("avg_package", ["میانگینوزنبسته"]),
    ("avg_net", ["میانگینوزنخالص", "میانگینحجمخالص"]),
    ("sd_net", ["انحرافاستانداردوزن", "انحرافاستانداردحجم", "انحرافمعیار"]),
    ("corrected_avg", ["میانگینوزنخالصتصحیح", "میانگینحجمخالصتصحیح"]),
    ("t1_fail_count", ["تعدادبستههایردوددرآزمونt1", "تعدادبستههایردوددرازمونt1"]),
    ("t2_fail_count", ["تعدادبستههایردوددرآزمونt2", "تعدادبستههایردوددرازمونt2"]),
    ("t1_result", ["نتیجهآزمونt1"]),
    ("t2_result", ["نتیجهآزمونt2"]),
    ("avg_requirement", ["آزمونالزاممیانگینبستهها"]),
    ("shortage", ["درصدکمبودوزن", "درصدکمبودحجم"]),
    ("overall_result", ["نتیجهکلیاستاندارد16381"]),
    ("declared_tolerance", ["رواداریاظهارشده"]),
]

# =========================
# توابع کمکی عمومی
# =========================
def norm_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).translate(PERSIAN_DIGITS)
    s = s.replace("ي", "ی").replace("ك", "ک")
    s = s.replace("\u200c", "")
    s = re.sub(r"[\s\-\–\—\(\)\[\]\{\}\.,:;،/\\]+", "", s)
    s = re.sub(r"[^\w\u0600-\u06FF]+", "", s)
    return s.lower()

def to_float(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        if math.isnan(float(v)):
            return None
        return float(v)
    s = str(v).translate(PERSIAN_DIGITS).strip()
    s = s.replace("ي", "ی").replace("ك", "ک").replace(",", ".")
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None

def fmt_num(x: Any, nd: int = 6) -> str:
    if x is None:
        return "نامشخص"
    try:
        x = float(x)
    except Exception:
        return str(x)
    s = f"{x:.{nd}f}".rstrip("0").rstrip(".")
    return s if s else "0"

def has_formula(cell) -> bool:
    return isinstance(cell.value, str) and cell.value.startswith("=")

def stdev_sample(xs: List[float]) -> Optional[float]:
    return statistics.stdev(xs) if len(xs) >= 2 else None

# =========================
# تشخیص ساختار فایل
# =========================
def find_header_row(ws, max_scan: int = 20) -> int:
    best_score, best_row = -1, 1
    keywords = [
        "وزنناخالص", "وزنبسته", "وزنخالص", "حجمخالص",
        "خطایt1", "خطایt2", "نامفراورده", "اطلاعاتونتایجآزمون",
    ]
    for r in range(1, min(ws.max_row, max_scan) + 1):
        joined = " ".join(norm_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 15) + 1))
        score = sum(1 for kw in keywords if kw in joined)
        if score > best_score:
            best_score, best_row = score, r
    return best_row

def find_column_map(ws, header_row: int) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for c in range(1, min(ws.max_column, 20) + 1):
        txt = norm_text(ws.cell(header_row, c).value)
        if not txt:
            continue
        for key, kws in COLUMN_KEYWORDS.items():
            if key in mapping:
                continue
            if any(kw in txt for kw in kws):
                mapping[key] = c
    return mapping

def find_sample_end(ws, header_row: int, col_map: Dict[str, int]) -> int:
    core_cols = [col_map[k] for k in ["gross", "package", "net_weight", "net_volume", "corrected_weight", "corrected_volume", "t1_error", "t2_error"] if k in col_map]
    if not core_cols:
        core_cols = list(range(1, min(ws.max_column, 8) + 1))
    last = header_row
    for r in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, "") for c in core_cols):
            last = r
    return last

def extract_series(ws, wsv, col_idx: int, start_row: int, end_row: int):
    values: List[float] = []
    row_details: List[Dict[str, Any]] = []
    for r in range(start_row, end_row + 1):
        cell_f = ws.cell(r, col_idx)
        cell_v = wsv.cell(r, col_idx)
        num = to_float(cell_v.value)
        row_details.append({
            "row": r,
            "coord": cell_f.coordinate,
            "value": cell_v.value,
            "formula": cell_f.value if has_formula(cell_f) else None,
            "is_formula": has_formula(cell_f),
            "numeric": num,
        })
        if num is not None:
            values.append(num)
    return values, row_details

def pick_primary_series(series_meta: Dict[str, Dict[str, Any]]) -> Optional[str]:
    priority = {"gross": 0, "package": 1, "net_weight": 2, "net_volume": 3, "corrected_weight": 4, "corrected_volume": 5}
    best = None
    for key, meta in series_meta.items():
        xs = meta["values"]
        if len(xs) < 5:
            continue
        sdv = stdev_sample(xs)
        score = (len(xs), sdv if sdv is not None else -1, -priority.get(key, 99))
        if best is None or score > best[0]:
            best = (score, key)
    return best[1] if best else None

def extract_nominal(sheet_name: str, summary_nominal=None, product_text=None, default_nominal=None):
    for source in [sheet_name, product_text or "", str(summary_nominal) if summary_nominal is not None else "", str(default_nominal) if default_nominal is not None else ""]:
        m = re.search(r"([0-9]+(?:\.[0-9]+)?)", str(source).translate(PERSIAN_DIGITS))
        if m:
            return float(m.group(0)), source
    return None, "نامشخص"

def get_correction_factor(n: int) -> float:
    if n >= 120:
        return 0.234
    elif n >= 80:
        return 0.295
    else:
        return 0.379

# =========================
# آزمون‌های آماری
# =========================
def runs_test(xs: List[float]):
    if len(xs) < 10:
        return None
    med = statistics.median(xs)
    signs = [1 if x > med else -1 if x < med else 0 for x in xs]
    signs = [s for s in signs if s != 0]
    n1 = sum(1 for s in signs if s == 1)
    n2 = sum(1 for s in signs if s == -1)
    if n1 == 0 or n2 == 0:
        return {"runs": None, "z": None, "p": None, "median": med, "n1": n1, "n2": n2, "used_n": len(signs)}
    runs = 1 + sum(1 for i in range(1, len(signs)) if signs[i] != signs[i-1])
    mu = 1 + (2 * n1 * n2) / (n1 + n2)
    var = (2 * n1 * n2 * (2 * n1 * n2 - n1 - n2)) / (((n1 + n2) ** 2) * (n1 + n2 - 1))
    z = (runs - mu) / math.sqrt(var) if var > 0 else None
    p = 2 * (1 - stats.norm.cdf(abs(z))) if z is not None else None
    return {"runs": runs, "z": z, "p": p, "median": med, "n1": n1, "n2": n2, "used_n": len(signs), "mu": mu}

def benford(xs: List[float]):
    digs = []
    for x in xs:
        a = abs(float(x))
        if a <= 0:
            continue
        s = re.sub(r"[^0-9]", "", f"{a:.15g}").lstrip("0")
        if s:
            d = int(s[0])
            if 1 <= d <= 9:
                digs.append(d)
    if len(digs) < 20:
        return None
    obs = Counter(digs)
    n = len(digs)
    exp = {d: math.log10(1 + 1 / d) * n for d in range(1, 10)}
    chi2 = sum((obs.get(d, 0) - exp[d]) ** 2 / exp[d] for d in range(1, 10))
    return {"n": n, "obs": obs, "chi2": chi2, "p": float(stats.chi2.sf(chi2, df=8))}

def duplicate_stats(xs: List[float]):
    c = Counter(xs)
    dup = sum(v - 1 for v in c.values() if v > 1)
    return {"unique": len(c), "dup_count": dup, "dup_rate": dup / len(xs) if xs else None, "most_common": c.most_common(5)}

def decimal_stats(xs: List[float]):
    if not xs:
        return None
    dig1, dig2 = [], []
    for x in xs:
        s1 = f"{float(x):.1f}"
        s2 = f"{float(x):.2f}"
        dig1.append(s1.split(".", 1)[1][-1])
        dig2.append(s2.split(".", 1)[1])
    c1, c2 = Counter(dig1), Counter(dig2)
    return {
        "n": len(xs),
        "pct_0": c1.get("0", 0) / len(dig1) * 100,
        "pct_5": c1.get("5", 0) / len(dig1) * 100,
        "pct_0_or_5": (c1.get("0", 0) + c1.get("5", 0)) / len(dig1) * 100,
        "unique_decimals_2dp": len(c2),
        "diversity_ratio": len(c2) / len(dig2),
        "top_2dp": c2.most_common(5),
    }

def lag1(xs: List[float]):
    if len(xs) < 3:
        return None
    a, b = xs[:-1], xs[1:]
    if len(set(a)) < 2 or len(set(b)) < 2:
        return {"r": None, "p": None}
    r, p = stats.pearsonr(a, b)
    return {"r": float(r), "p": float(p)}

def iqr_outliers(xs: List[float]):
    if len(xs) < 4:
        return None
    q1, q3 = statistics.quantiles(xs, n=4, method="inclusive")[0], statistics.quantiles(xs, n=4, method="inclusive")[2]
    iqr = q3 - q1
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    idx = [i for i, x in enumerate(xs) if x < lo or x > hi]
    return {"q1": q1, "q3": q3, "iqr": iqr, "low": lo, "high": hi, "count": len(idx), "indices": idx}

def shapiro(xs: List[float]):
    if len(xs) < 3 or len(xs) > 5000:
        return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        w, p = stats.shapiro(xs)
    return {"W": float(w), "p": float(p)}

def extreme_outliers(xs: List[float], n_sigma: float = 3.0) -> int:
    if len(xs) < 2:
        return 0
    mean = sum(xs) / len(xs)
    sd = stdev_sample(xs)
    if sd is None or sd == 0:
        return 0
    return sum(1 for x in xs if abs(x - mean) > n_sigma * sd)

def get_duplicate_details(xs: List[float], max_items: int = 5) -> str:
    cnt = Counter(xs)
    duplicates = [(val, count) for val, count in cnt.items() if count > 1]
    if not duplicates:
        return "هیچ مقدار تکراری (بیش از یک بار) یافت نشد."
    duplicates.sort(key=lambda x: x[1], reverse=True)
    limited = duplicates[:max_items]
    parts = [f"مقدار {fmt_num(val, 3)}: {count} بار" for val, count in limited]
    if len(duplicates) > max_items:
        parts.append(f"و {len(duplicates) - max_items} مورد دیگر ...")
    return "؛ ".join(parts)

def detect_patterns(xs: List[float], dec_stats: Optional[Dict]) -> str:
    patterns = []
    if len(xs) >= 4:
        diffs = [xs[i+1] - xs[i] for i in range(len(xs)-1)]
        if len(set(round(d, 4) for d in diffs)) == 1:
            step = diffs[0]
            patterns.append(f"روند خطی با گام ثابت {step:.4f}")
    block_repeat = 0
    i = 0
    while i < len(xs) - 1:
        if xs[i] == xs[i+1]:
            block_repeat += 1
            i += 2
        else:
            i += 1
    if block_repeat >= 3:
        patterns.append(f"تکرار بلوکی (جفت‌های تکراری متوالی) در {block_repeat} جفت")
    if dec_stats and dec_stats.get("pct_0_or_5", 0) > 50:
        patterns.append(f"گرد شدن بیش از حد: {dec_stats['pct_0_or_5']:.1f}% اعداد رقم دهم 0 یا 5 دارند")
    max_run = 0
    cur_run = 1
    for i in range(1, len(xs)):
        if xs[i] == xs[i-1]:
            cur_run += 1
        else:
            max_run = max(max_run, cur_run)
            cur_run = 1
    max_run = max(max_run, cur_run)
    if max_run >= 5:
        patterns.append(f"یک دنباله تکراری به طول {max_run} ردیف پشت سر هم وجود دارد")
    return "؛ ".join(patterns) if patterns else "الگوی خاصی تشخیص داده نشد."

# =========================
# خلاصه و محاسبات تطبیقی
# =========================
def find_summary(ws, wsv):
    found = {}
    for r in range(1, min(ws.max_row, 40) + 1):
        for c in range(1, min(ws.max_column, 12) + 1):
            label = ws.cell(r, c).value
            keytxt = norm_text(label)
            if not keytxt:
                continue
            for key, kws in SUMMARY_PATTERNS:
                if key in found:
                    continue
                if any(kw in keytxt for kw in kws):
                    value_col = c - 1 if c > 1 else None
                    found[key] = {
                        "label": label,
                        "label_cell": ws.cell(r, c).coordinate,
                        "value_cell": ws.cell(r, value_col).coordinate if value_col else None,
                        "value": wsv.cell(r, value_col).value if value_col else None,
                        "formula": ws.cell(r, value_col).value if value_col else None,
                        "row": r,
                        "col": value_col,
                    }
    return found

def detect_measure_type(series_meta: Dict[str, Dict[str, Any]], summary: Dict) -> str:
    if "net_volume" in series_meta and series_meta["net_volume"]["count"] > 0:
        return "حجم"
    if "net_weight" in series_meta and series_meta["net_weight"]["count"] > 0:
        return "وزن"
    for key, val in summary.items():
        if "وزن" in str(val.get("label", "")):
            return "وزن"
        if "حجم" in str(val.get("label", "")):
            return "حجم"
    return "وزن"

def expected_summary(series_meta: Dict[str, Dict[str, Any]], nominal: Optional[float], measure_type: str, sample_count: int) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    package = series_meta.get("package", {}).get("values", [])
    if measure_type == "وزن":
        net_key = "net_weight"
    else:
        net_key = "net_volume"
    net_vals = series_meta.get(net_key, {}).get("values", [])
    t1 = series_meta.get("t1_error", {}).get("values", [])
    t2 = series_meta.get("t2_error", {}).get("values", [])
    cf = get_correction_factor(sample_count)
    if package:
        out["avg_package"] = sum(package) / len(package)
    if net_vals:
        out["avg_net"] = sum(net_vals) / len(net_vals)
        if len(net_vals) >= 2:
            out["sd_net"] = statistics.stdev(net_vals)
            out["corrected_avg"] = cf * out["sd_net"] + out["avg_net"]
    if t1:
        out["t1_fail_count"] = sum(1 for x in t1 if to_float(x) not in (None, 0.0))
        out["t1_result"] = "قبول" if out["t1_fail_count"] < 8 else "مردود"
    if t2:
        out["t2_fail_count"] = sum(1 for x in t2 if to_float(x) not in (None, 0.0))
        out["t2_result"] = "قبول" if out["t2_fail_count"] < 1 else "مردود"
    if "corrected_avg" in out and nominal is not None:
        out["avg_requirement"] = "قبول" if out["corrected_avg"] >= nominal else "مردود"
        if "avg_net" in out:
            out["shortage"] = 0 if out["avg_requirement"] == "قبول" else (nominal - out["avg_net"]) / nominal * 100
    if all(k in out for k in ["t1_result", "t2_result", "avg_requirement"]):
        out["overall_result"] = (out["t1_result"] == "قبول" and out["t2_result"] == "قبول" and out["avg_requirement"] == "قبول")
    return out

# =========================
# پردازش اصلی هر فایل
# =========================
def process_file(path: Path, global_wb: Workbook, all_records: List[Dict]):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=False)
    wv = load_workbook(path, data_only=True)

    for sname in wb.sheetnames:
        ws = wb[sname]
        dvs = wv[sname]

        header_row = find_header_row(ws)
        col_map = find_column_map(ws, header_row)
        sample_end = find_sample_end(ws, header_row, col_map)
        sample_start = header_row + 1
        sample_count = max(0, sample_end - sample_start + 1)

        series_meta = {}
        for key in ["gross", "package", "net_weight", "net_volume", "corrected_weight", "corrected_volume", "t1_error", "t2_error"]:
            if key in col_map:
                xs, details = extract_series(ws, dvs, col_map[key], sample_start, sample_end)
                series_meta[key] = {"values": xs, "details": details, "count": len(xs), "col": col_map[key]}

        primary = pick_primary_series(series_meta)
        primary_values = series_meta[primary]["values"] if primary else []

        stats_primary = {"n": len(primary_values), "mean": None, "sd": None, "cv": None}
        if primary_values:
            stats_primary["mean"] = sum(primary_values) / len(primary_values)
            stats_primary["sd"] = stdev_sample(primary_values)
            if stats_primary["mean"] not in (None, 0) and stats_primary["sd"] is not None:
                stats_primary["cv"] = stats_primary["sd"] / stats_primary["mean"] * 100

        run = runs_test(primary_values) if primary_values else None
        ben = benford(primary_values) if primary_values else None
        dup = duplicate_stats(primary_values) if primary_values else None
        dec = decimal_stats(primary_values) if primary_values else None
        lag = lag1(primary_values) if primary_values else None
        out_iqr = iqr_outliers(primary_values) if primary_values else None
        shp = shapiro(primary_values) if primary_values else None
        extreme_cnt = extreme_outliers(primary_values, 3.0) if primary_values else 0

        duplicate_detail = get_duplicate_details(primary_values, 5) if primary_values else "نامشخص"
        pattern_detail = detect_patterns(primary_values, dec) if primary_values else "نامشخص"

        summary = find_summary(ws, dvs)
        product_text = dvs["A2"].value if sample_start <= 2 <= sample_end else None
        nominal, nominal_source = extract_nominal(
            ws.title,
            summary.get("nominal", {}).get("value") if "nominal" in summary else None,
            product_text,
            default_nominal=DEFAULT_NOMINAL,
        )
        measure_type = detect_measure_type(series_meta, summary)
        exp = expected_summary(series_meta, nominal, measure_type, sample_count)

        reasons = []
        risk = 0
        critical = False

        if "gross" not in col_map:
            critical = True
            reasons.append("ستون وزن ناخالص شناسایی نشد.")
            risk += 25
        else:
            if series_meta.get("gross", {}).get("count", 0) == 0:
                critical = True
                reasons.append("ستون وزن ناخالص پیدا شد اما هیچ مقدار عددی از آن استخراج نشد.")
                risk += 25

        net_key = "net_weight" if measure_type == "وزن" else "net_volume"
        if net_key not in col_map:
            critical = True
            reasons.append(f"ستون {('وزن خالص' if measure_type=='وزن' else 'حجم خالص')} پیدا نشد.")
            risk += 30
        else:
            net_count = series_meta.get(net_key, {}).get("count", 0)
            if net_count < sample_count:
                reasons.append(f"ستون {('وزن خالص' if measure_type=='وزن' else 'حجم خالص')} فقط {net_count} مقدار عددی از {sample_count} ردیف دارد.")

        for key, label in [("t1_error", "خطای T1"), ("t2_error", "خطای T2")]:
            if key not in col_map:
                reasons.append(f"ستون {label} پیدا نشد.")
                if key in ("t1_error", "t2_error"):
                    critical = True

        for key, label in [(net_key, "مقدار خالص"), ("t1_error", "خطای T1"), ("t2_error", "خطای T2")]:
            if key not in col_map:
                continue
            c = col_map[key]
            formula_rows = [r for r in range(sample_start, sample_end+1) if to_float(dvs.cell(r, c).value) is not None and has_formula(ws.cell(r, c))]
            coverage = len(formula_rows) / sample_count if sample_count else 0
            if coverage < 0.95:
                reasons.append(f"فرمول ستون {label} در {len(formula_rows)}/{sample_count} ردیف باقی مانده است.")
                risk += int((1 - coverage) * (35 if key == net_key else 25))

        def check_summary(key, calc, label, tol=1e-6):
            nonlocal risk
            if key in summary and calc is not None:
                rec = to_float(summary[key]["value"])
                if rec is not None and abs(rec - calc) > tol:
                    reasons.append(f"{label} ثبت‌شده ({fmt_num(rec,6)}) با محاسبه مستقل ({fmt_num(calc,6)}) هم‌خوان نیست.")
                    rel_diff = abs(rec - calc) / max(abs(calc), 1e-6)
                    risk += 20 if rel_diff > 0.05 else 10

        if "avg_package" in exp:
            check_summary("avg_package", exp["avg_package"], "میانگین وزن بسته‌ها")
        if "avg_net" in exp:
            check_summary("avg_net", exp["avg_net"], f"میانگین {measure_type} خالص")
        if "sd_net" in exp:
            check_summary("sd_net", exp["sd_net"], f"انحراف معیار {measure_type} خالص")
        if "corrected_avg" in exp:
            check_summary("corrected_avg", exp["corrected_avg"], f"میانگین {measure_type} خالص تصحیح‌شده")

        overall = None
        if "overall_result" in summary:
            ov = summary["overall_result"]["value"]
            overall = bool(ov) if isinstance(ov, bool) else str(ov).strip().lower() in ("true", "1", "قبول", "yes")
        if overall is False:
            reasons.append("نتیجه کلی فرم False/مردود است.")
            risk += 10

        if extreme_cnt > 0:
            reasons.append(f"{extreme_cnt} داده با فاصله بیش از ۳ انحراف معیار از میانگین (پرت شدید) شناسایی شد.")
            risk += min(30, extreme_cnt * 15)

        if nominal and "avg_net" in exp and exp["avg_net"] is not None:
            if abs(exp["avg_net"] - nominal) / nominal > 0.10:
                reasons.append(f"میانگین {measure_type} خالص ({fmt_num(exp['avg_net'],2)}) نسبت به مقدار اسمی ({nominal}) بیش از 10% اختلاف دارد.")
                risk += 15

        if primary_values:
            if dup and dup["dup_rate"] and dup["dup_rate"] > 0.3:
                reasons.append(f"نرخ تکرار {dup['dup_rate']*100:.1f}% ({dup['dup_count']} تکرار افزوده).")
                risk += 8
            if run and run.get("p") and run["p"] < 0.05:
                reasons.append(f"Runs Test معنادار (p={run['p']:.4g})")
                risk += 5
            if shp and shp.get("p") and shp["p"] < 0.01:
                reasons.append(f"Shapiro-Wilk نرمال نیست (p={shp['p']:.4g})")
                risk += 5
            if out_iqr and out_iqr["count"] > 0:
                reasons.append(f"{out_iqr['count']} داده پرت IQR")
                risk += min(10, out_iqr["count"] * 3)
            if dec and dec["pct_0_or_5"] > 50:
                reasons.append(f"{dec['pct_0_or_5']:.1f}% رقم دهم 0/5")
                risk += 5
            if ben and ben.get("p") and ben["p"] < 0.001:
                reasons.append("Benford p بسیار کوچک")
                risk += 2

        if primary is None:
            critical = True
            reasons.append("هیچ ستون عددی قابل‌اعتماد یافت نشد.")
            risk += 30
        if sample_count < MIN_ROWS_FOR_ANALYSIS:
            critical = True
            reasons.append("تعداد نمونه کافی نیست.")
            risk += 20

        if critical:
            result = "خطا"
        else:
            result = "بسیار مشکوک" if risk >= 65 else "مشکوک" if risk >= 35 else "سالم"

        row = {
            "نام فایل": path.name,
            "نام شیت": sname,
            "تعداد نمونه": sample_count,
            "مقدار اسمی": nominal if nominal is not None else "نامشخص",
            "منبع مقدار اسمی": nominal_source,
            "نوع اندازه‌گیری": measure_type,
            "نتیجه نهایی": result,
            "امتیاز ریسک": min(100, int(round(risk))),
            "دلایل دقیق و کامل": "؛ ".join(reasons) if reasons else "مورد خاصی مشاهده نشد.",
            "جزئیات تکرار و الگو": f"تکرارها: {duplicate_detail} | الگوها: {pattern_detail}",
            "ستون اصلی تحلیل": primary if primary else "نامشخص",
            "میانگین محاسبه‌شده": stats_primary["mean"],
            "SD محاسبه‌شده": stats_primary["sd"],
            "CV%": stats_primary["cv"],
            "Runs": run["runs"] if run else None,
            "Runs z": run["z"] if run else None,
            "Runs p": run["p"] if run else None,
            "Benford p": ben["p"] if ben else None,
            "Lag-1 r": lag["r"] if lag else None,
            "Lag-1 p": lag["p"] if lag else None,
            "تکرارهای افزوده": dup["dup_count"] if dup else None,
            "نرخ تکرار%": dup["dup_rate"]*100 if dup and dup["dup_rate"] else None,
            "درصد رقم‌های دهم 0/5": dec["pct_0_or_5"] if dec else None,
            "تنوع اعشار 2dp": dec["diversity_ratio"] if dec else None,
            "IQR outlier": out_iqr["count"] if out_iqr else 0,
            "Shapiro p": shp["p"] if shp else None,
            "پوشش فرمول مقدار خالص": None if net_key not in col_map else len([r for r in range(sample_start, sample_end+1) if has_formula(ws.cell(r, col_map[net_key])) and to_float(dvs.cell(r, col_map[net_key]).value) is not None]) / sample_count if sample_count else None,
            "پوشش فرمول خطای T1": None if "t1_error" not in col_map else len([r for r in range(sample_start, sample_end+1) if has_formula(ws.cell(r, col_map["t1_error"])) and to_float(dvs.cell(r, col_map["t1_error"]).value) is not None]) / sample_count if sample_count else None,
            "پوشش فرمول خطای T2": None if "t2_error" not in col_map else len([r for r in range(sample_start, sample_end+1) if has_formula(ws.cell(r, col_map["t2_error"])) and to_float(dvs.cell(r, col_map["t2_error"]).value) is not None]) / sample_count if sample_count else None,
            "تعداد داده عددی وزن ناخالص": series_meta.get("gross", {}).get("count") if "gross" in series_meta else None,
            "تعداد داده عددی وزن بسته": series_meta.get("package", {}).get("count") if "package" in series_meta else None,
            "تعداد داده عددی وزن خالص": series_meta.get("net_weight", {}).get("count") if "net_weight" in series_meta else None,
            "تعداد داده عددی حجم خالص": series_meta.get("net_volume", {}).get("count") if "net_volume" in series_meta else None,
            "نتیجه کلی فرم": "قبول" if overall is True else "مردود" if overall is False else "نامشخص",
            "تعداد برچسب‌های خلاصه شناسایی‌شده": len(summary),
        }
        all_records.append(row)

        # ذخیره داده‌های خام در شیت جداگانه
        if primary_values and len(primary_values) >= 5:
            data_sheet_name = f"داده_{path.stem[:20]}_{sname[:15]}"
            if data_sheet_name in global_wb.sheetnames:
                data_sheet = global_wb[data_sheet_name]
            else:
                data_sheet = global_wb.create_sheet(title=data_sheet_name)
            data_sheet["A1"] = "ردیف"
            data_sheet["B1"] = "مقدار"
            for idx, val in enumerate(primary_values, start=2):
                data_sheet.cell(row=idx, column=1, value=idx-1)
                data_sheet.cell(row=idx, column=2, value=val)
            data_sheet["D1"] = "برای رسم نمودار در اکسل:"
            data_sheet["D2"] = "1- محدوده داده‌ها (ستون‌های A و B) را انتخاب کنید."
            data_sheet["D3"] = "2- از سربرگ Insert → Charts نوع نمودار دلخواه (مثلاً هیستوگرام) را انتخاب کنید."
    return all_records

def build_final_report(records: List[Dict], output_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "گزارش خلاصه"

    headers = [
        "نام فایل", "نام شیت", "تعداد نمونه", "مقدار اسمی", "منبع مقدار اسمی",
        "نوع اندازه‌گیری", "نتیجه نهایی", "امتیاز ریسک", "دلایل دقیق و کامل",
        "جزئیات تکرار و الگو", "ستون اصلی تحلیل", "میانگین محاسبه‌شده", "SD محاسبه‌شده", "CV%",
        "Runs", "Runs z", "Runs p", "Benford p", "Lag-1 r", "Lag-1 p",
        "تکرارهای افزوده", "نرخ تکرار%", "درصد رقم‌های دهم 0/5", "تنوع اعشار 2dp",
        "IQR outlier", "Shapiro p", "پوشش فرمول مقدار خالص", "پوشش فرمول خطای T1",
        "پوشش فرمول خطای T2", "تعداد داده عددی وزن ناخالص", "تعداد داده عددی وزن بسته",
        "تعداد داده عددی وزن خالص", "تعداد داده عددی حجم خالص", "نتیجه کلی فرم",
        "تعداد برچسب‌های خلاصه شناسایی‌شده"
    ]
    ws.append(headers)
    for rec in records:
        ws.append([rec.get(h) for h in headers])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    green_fill = PatternFill("solid", fgColor="C6EFCE")
    orange_fill = PatternFill("solid", fgColor="FFEB9C")
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        result = row[6].value
        fill_color = None
        if result == "سالم":
            fill_color = green_fill
        elif result == "مشکوک":
            fill_color = orange_fill
        elif result in ("بسیار مشکوک", "خطا"):
            fill_color = red_fill
        if fill_color:
            for cell in row:
                cell.fill = fill_color

    widths = {1:28,2:18,3:12,4:14,5:30,6:14,7:16,8:12,9:60,10:50,11:16,12:16,13:14,14:10,
              15:12,16:12,17:12,18:12,19:12,20:12,21:12,22:14,23:14,24:12,25:12,26:14,
              27:14,28:14,29:14,30:16,31:16,32:16,33:16,34:12,35:16}
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws.freeze_panes = "A2"
    ws.sheet_view.rightToLeft = True
    ws.auto_filter.ref = ws.dimensions
    wb.save(output_path)

# =========================
# رابط کاربری Streamlit
# =========================
st.set_page_config(page_title="بازرسی فایل‌های اکسل", layout="wide")
st.title("📋 بازرسی پیشرفته فایل‌های اکسل")

st.markdown("""
فایل‌های اکسل خود را آپلود کنید تا از نظر سلامت داده‌ها، فرمول‌ها و نشانه‌های داده‌سازی مصنوعی بررسی شوند.
""")

uploaded_files = st.file_uploader(
    "فایل‌های اکسل (xlsx.) را انتخاب کنید",
    type="xlsx",
    accept_multiple_files=True,
    help="می‌توانید چند فایل را با هم انتخاب کنید یا بکشید و رها کنید."
)

if st.button("🚀 شروع پردازش"):
    if not uploaded_files:
        st.warning("لطفاً حداقل یک فایل انتخاب کنید.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()

        global_wb = Workbook()
        global_wb.remove(global_wb.active)
        all_records = []

        total = len(uploaded_files)

        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.info(f"در حال پردازش: {uploaded_file.name} ...")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = Path(tmp.name)

            try:
                process_file(tmp_path, global_wb, all_records)
            except Exception as e:
                all_records.append({
                    "نام فایل": uploaded_file.name,
                    "نام شیت": "خطا",
                    "نتیجه نهایی": "خطا",
                    "امتیاز ریسک": 100,
                    "دلایل دقیق و کامل": f"خطا در پردازش: {e}",
                    "جزئیات تکرار و الگو": "نامشخص",
                })

            tmp_path.unlink()
            progress_bar.progress((idx + 1) / total)

        status_text.success("✅ پردازش همه فایل‌ها به پایان رسید.")

        df = pd.DataFrame(all_records)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📁 فایل‌های پردازش‌شده", len(df))
        with col2:
            st.metric("📄 تعداد شیت‌ها", len(df))
        if "نتیجه نهایی" in df.columns:
            with col3:
                counts = df["نتیجه نهایی"].value_counts()
                st.metric("🟢 سالم", counts.get("سالم", 0))
                st.metric("🟠 مشکوک", counts.get("مشکوک", 0))
                st.metric("🔴 بسیار مشکوک/خطا", counts.get("بسیار مشکوک",0)+counts.get("خطا",0))

        st.subheader("📊 جدول کامل نتایج")
        st.dataframe(df, use_container_width=True, hide_index=True)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
            build_final_report(all_records, Path(tmp_out.name))
            with open(tmp_out.name, "rb") as f:
                st.download_button(
                    label="📥 دانلود گزارش نهایی (Excel)",
                    data=f.read(),
                    file_name="گزارش_نهایی_پایدار.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )