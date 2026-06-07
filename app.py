# -*- coding: utf-8 -*-
"""
app.py
رابط کاربری Streamlit برای موتور بازرسی پیش‌بسته‌بندی

فایل‌های لازم کنار هم:
    app.py
    inspection_engine.py

اجرا:
    streamlit run app.py
"""

import inspect
import tempfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

from inspection_engine import (
    Config,
    RESULT_LABELS,
    process_workbook,
    error_record,
    build_report,
)


# ═════════════════════════════════════════════════════════════
# تنظیمات صفحه
# ═════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="سامانه بازرسی پیش‌بسته‌بندی",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ═════════════════════════════════════════════════════════════
# ابزارهای عمومی
# ═════════════════════════════════════════════════════════════

L = RESULT_LABELS


def label(key, fallback):
    """
    گرفتن عنوان فارسی ستون از RESULT_LABELS.
    اگر کلید وجود نداشت، مقدار جایگزین برمی‌گردد.
    """
    return L.get(key, fallback)


COL = {
    "file": label("file", "نام فایل"),
    "sheet": label("sheet", "نام شیت"),
    "n": label("n", "تعداد نمونه"),
    "mtype": label("mtype", "نوع اندازه‌گیری"),
    "unit": label("unit", "واحد"),
    "nominal": label("nominal", "مقدار اسمی"),
    "form_result": label("form_result", "نتیجه درج‌شده در فرم"),
    "standard_result": label("standard_result", "نتیجه استاندارد"),
    "file_verdict": label("file_verdict", "اعتبارسنجی فایل"),
    "risk": label("risk", "امتیاز ریسک"),
    "standard_reasons": label("standard_reasons", "دلایل استاندارد"),
    "file_reasons": label("file_reasons", "دلایل اعتبارسنجی فایل"),
    "tne": label("tne", "حد خطای مجاز TNE"),
    "tne_pct": label("tne_pct", "درصد TNE"),
    "form_tolerance": label("form_tolerance", "رواداری فرم"),
    "mean": label("mean", "میانگین"),
    "sd": label("sd", "انحراف معیار"),
    "corrected_avg": label("corrected_avg", "میانگین اصلاح‌شده"),
    "cv": label("cv", "ضریب تغییرات CV%"),
    "t1_fail": label("t1_fail", "تعداد مردودی T1"),
    "t1_allowed": label("t1_allowed", "مجاز T1"),
    "t2_fail": label("t2_fail", "تعداد مردودی T2"),
    "avg_req": label("avg_req", "الزام میانگین"),
    "t1_req": label("t1_req", "الزام T1"),
    "t2_req": label("t2_req", "الزام T2"),
    "net_calc_errors": label("net_calc_errors", "خطاهای محاسبه خالص"),
    "shortage_errors": label("shortage_errors", "خطاهای درصد کمبود"),
    "dup_rate": label("dup_rate", "نرخ تکرار"),
    "blocks": label("blocks", "بلوک‌های تکراری"),
    "iqr": label("iqr", "داده پرت IQR"),
    "ex3": label("ex3", "داده فراتر از ۳σ"),
    "runs_p": label("runs_p", "Runs Test p-value"),
    "round_pct": label("round_pct", "درصد اعداد رند"),
}


def safe_get(record, key, default="—"):
    try:
        value = record.get(key, default)
        if value is None or value == "":
            return default
        return value
    except Exception:
        return default


def as_int(value, default=0):
    try:
        if value in (None, "", "—"):
            return default
        return int(float(str(value).replace(",", "")))
    except Exception:
        return default


def as_float(value, default=0.0):
    try:
        if value in (None, "", "—"):
            return default
        return float(str(value).replace(",", ""))
    except Exception:
        return default


def bounded_risk(value):
    r = as_int(value, 0)
    return max(0, min(100, r))


def risk_level(risk):
    risk = bounded_risk(risk)

    if risk >= 65:
        return "بالا"

    if risk >= 35:
        return "متوسط"

    return "پایین"


def split_reasons(text):
    if text in (None, "", "—"):
        return []

    return [x.strip() for x in str(text).split("|") if x.strip()]


def rerun_app():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


def card():
    """
    کانتینر کارت‌مانند.
    اگر نسخه Streamlit قدیمی باشد و border را پشتیبانی نکند،
    به کانتینر ساده برمی‌گردد.
    """
    try:
        return st.container(border=True)
    except TypeError:
        return st.container()


def show_status(value):
    """
    نمایش وضعیت بدون HTML خام.
    """
    value = str(value)

    if value in ("قبول", "سالم"):
        st.success(value)
    elif value in ("مشکوک",):
        st.warning(value)
    elif value in ("مردود", "بسیار مشکوک", "خطا"):
        st.error(value)
    elif value in ("نامشخص", "—"):
        st.info(value)
    else:
        st.info(value)


def show_risk(value):
    """
    نمایش ریسک بدون HTML خام.
    """
    risk = bounded_risk(value)
    level = risk_level(risk)

    if level == "بالا":
        st.error(f"ریسک بالا: {risk}")
    elif level == "متوسط":
        st.warning(f"ریسک متوسط: {risk}")
    else:
        st.success(f"ریسک پایین: {risk}")

    st.progress(risk)


def show_metric_card(title, value, help_text=None):
    with card():
        st.metric(title, value)
        if help_text:
            st.caption(help_text)


def count_by(records, column):
    out = {}

    for r in records:
        v = safe_get(r, column, "نامشخص")
        out[str(v)] = out.get(str(v), 0) + 1

    return out


def records_to_df(records, columns=None):
    if not records:
        return pd.DataFrame()

    if columns is None:
        all_cols = []
        for r in records:
            for k in r.keys():
                if k not in all_cols:
                    all_cols.append(k)
        columns = all_cols

    rows = []

    for r in records:
        rows.append({c: safe_get(r, c) for c in columns})

    return pd.DataFrame(rows)


def show_count_chart(title, counts):
    """
    نمودار شمارشی کاملاً Native.
    هیچ HTML تولید نمی‌کند.
    """
    with card():
        st.subheader(title)

        if not counts:
            st.info("داده‌ای برای نمایش وجود ندارد.")
            return

        df = pd.DataFrame(
            [
                {"وضعیت": str(k), "تعداد": int(v)}
                for k, v in counts.items()
            ]
        )

        st.bar_chart(
            df.set_index("وضعیت"),
            use_container_width=True,
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )


def show_count_progress(title, counts):
    """
    نمایش شمارش‌ها به صورت Progress بدون HTML.
    """
    with card():
        st.subheader(title)

        if not counts:
            st.info("داده‌ای برای نمایش وجود ندارد.")
            return

        max_value = max(counts.values()) if counts else 1

        for k, v in counts.items():
            c1, c2 = st.columns([3, 1])

            with c1:
                st.write(str(k))

            with c2:
                st.write(f"**{v}**")

            percent = int((v / max_value) * 100) if max_value else 0
            st.progress(percent)


def show_detail_item(title, value):
    with card():
        st.caption(title)
        st.write(value)


def show_reason_box(title, reasons):
    with card():
        st.subheader(title)

        if reasons:
            for reason in reasons:
                st.write(f"• {reason}")
        else:
            st.success("موردی ثبت نشده است.")


def make_config(
    min_rows,
    near_dup_rel,
    net_abs_tol,
    percent_abs_tol,
    tolerance_rel_tol,
    strict_formula_coverage,
):
    """
    ساخت Config به شکل مقاوم.
    اگر در نسخه موتور بعضی پارامترها وجود نداشته باشند، حذف می‌شوند.
    """
    desired = {
        "min_rows": int(min_rows),
        "near_dup_rel": float(near_dup_rel),
        "net_abs_tol": float(net_abs_tol),
        "percent_abs_tol": float(percent_abs_tol),
        "tolerance_rel_tol": float(tolerance_rel_tol),
        "strict_formula_coverage": bool(strict_formula_coverage),
    }

    try:
        sig = inspect.signature(Config)
        allowed = set(sig.parameters.keys())
        kwargs = {k: v for k, v in desired.items() if k in allowed}
        return Config(**kwargs)
    except Exception:
        try:
            return Config()
        except Exception:
            return Config


def save_uploaded_files(uploaded_files, folder):
    paths = []

    for i, uploaded in enumerate(uploaded_files, start=1):
        suffix = Path(uploaded.name).suffix.lower()

        if suffix not in [".xlsx", ".xlsm"]:
            continue

        original_stem = Path(uploaded.name).stem
        original_suffix = Path(uploaded.name).suffix
        safe_name = f"{i:03d}_{original_stem}{original_suffix}"
        file_path = Path(folder) / safe_name

        with open(file_path, "wb") as f:
            f.write(uploaded.getbuffer())

        paths.append(file_path)

    return paths


def run_engine(paths, cfg):
    records = []
    total = len(paths)

    progress = st.progress(0)
    status = st.empty()

    for i, path in enumerate(paths, start=1):
        status.info(f"در حال پردازش فایل {i} از {total}: {path.name}")

        try:
            result = process_workbook(path, cfg)
            if result:
                records.extend(result)
        except Exception as ex:
            try:
                records.append(error_record(path.name, str(ex)))
            except Exception:
                records.append(
                    {
                        COL["file"]: path.name,
                        COL["sheet"]: "—",
                        COL["standard_result"]: "نامشخص",
                        COL["file_verdict"]: "خطا",
                        COL["risk"]: 100,
                        COL["file_reasons"]: str(ex),
                    }
                )

        progress.progress(i / total)

    status.success(f"پردازش کامل شد. تعداد رکوردهای تولیدشده: {len(records)}")
    progress.empty()

    return records


def build_excel_bytes(records):
    with tempfile.TemporaryDirectory() as td:
        out_path = Path(td) / "inspection_report.xlsx"
        build_report(records, out_path)

        with open(out_path, "rb") as f:
            return f.read()


def dataframe_with_risk(df):
    """
    نمایش DataFrame با ستون ریسک به صورت ProgressColumn در صورت امکان.
    """
    if df.empty:
        st.info("داده‌ای برای نمایش وجود ندارد.")
        return

    column_config = {}

    if COL["risk"] in df.columns:
        df[COL["risk"]] = pd.to_numeric(
            df[COL["risk"]],
            errors="coerce",
        ).fillna(0).astype(int)

        try:
            column_config[COL["risk"]] = st.column_config.ProgressColumn(
                COL["risk"],
                min_value=0,
                max_value=100,
                format="%d",
            )
        except Exception:
            column_config = {}

    try:
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            column_config=column_config,
        )
    except Exception:
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )


# ═════════════════════════════════════════════════════════════
# Session State
# ═════════════════════════════════════════════════════════════

if "records" not in st.session_state:
    st.session_state.records = []

if "last_run" not in st.session_state:
    st.session_state.last_run = None

if "last_files_count" not in st.session_state:
    st.session_state.last_files_count = 0


# ═════════════════════════════════════════════════════════════
# Sidebar
# ═════════════════════════════════════════════════════════════

with st.sidebar:
    st.title("⚙️ تنظیمات")
    st.caption("پارامترهای این بخش به موتور محاسباتی ارسال می‌شوند.")

    st.divider()

    min_rows = st.number_input(
        "حداقل تعداد نمونه",
        min_value=1,
        max_value=1000,
        value=20,
        step=1,
    )

    net_abs_tol = st.number_input(
        "تلورانس اختلاف مقدار خالص",
        min_value=0.0,
        max_value=100.0,
        value=0.01,
        step=0.01,
        format="%.4f",
    )

    percent_abs_tol = st.number_input(
        "تلورانس درصد کمبود",
        min_value=0.0,
        max_value=100.0,
        value=0.05,
        step=0.01,
        format="%.4f",
    )

    tolerance_rel_tol = st.number_input(
        "تلورانس نسبی رواداری",
        min_value=0.0,
        max_value=1.0,
        value=0.01,
        step=0.005,
        format="%.4f",
    )

    strict_formula_coverage = st.toggle(
        "سخت‌گیری در پوشش فرمول‌ها",
        value=False,
    )

    near_dup_rel = st.number_input(
        "پارامتر تشابه نزدیک",
        min_value=0.0,
        max_value=1.0,
        value=0.02,
        step=0.005,
        format="%.4f",
    )

    st.divider()

    if st.button("🧹 پاک‌کردن نتایج", use_container_width=True):
        st.session_state.records = []
        st.session_state.last_run = None
        st.session_state.last_files_count = 0
        rerun_app()

    st.divider()

    st.caption("فایل‌های مجاز: xlsx و xlsm")
    st.caption("خروجی: گزارش Excel")


# ═════════════════════════════════════════════════════════════
# Header
# ═════════════════════════════════════════════════════════════

st.title("🧪 سامانه هوشمند بازرسی پیش‌بسته‌بندی")

st.write(
    """
این سامانه فایل‌های Excel محصولات را پردازش می‌کند و نتیجه استاندارد،
اعتبارسنجی فایل، نتیجه فرم، شاخص‌های آماری و امتیاز ریسک را نمایش می‌دهد.
"""
)

st.caption("سازگار با موتور محاسباتی جدید inspection_engine.py")


# ═════════════════════════════════════════════════════════════
# Upload and Run
# ═════════════════════════════════════════════════════════════

with card():
    st.subheader("📥 بارگذاری فایل‌های Excel")

    uploaded_files = st.file_uploader(
        "فایل‌های اکسل را انتخاب کنید",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

    c1, c2, c3 = st.columns([1, 1, 2])

    with c1:
        run_btn = st.button(
            "🚀 شروع پردازش",
            type="primary",
            use_container_width=True,
            disabled=not uploaded_files,
        )

    with c2:
        if uploaded_files:
            st.metric("تعداد فایل انتخاب‌شده", len(uploaded_files))
        else:
            st.metric("تعداد فایل انتخاب‌شده", 0)

    with c3:
        st.caption(
            "پس از انتخاب فایل‌ها، روی دکمه شروع پردازش بزنید. "
            "هر شیت به صورت مستقل بررسی می‌شود."
        )


if run_btn:
    cfg = make_config(
        min_rows=min_rows,
        near_dup_rel=near_dup_rel,
        net_abs_tol=net_abs_tol,
        percent_abs_tol=percent_abs_tol,
        tolerance_rel_tol=tolerance_rel_tol,
        strict_formula_coverage=strict_formula_coverage,
    )

    with tempfile.TemporaryDirectory() as td:
        paths = save_uploaded_files(uploaded_files, td)

        if not paths:
            st.error("هیچ فایل Excel معتبری برای پردازش انتخاب نشده است.")
        else:
            with st.spinner("در حال پردازش فایل‌ها..."):
                st.session_state.records = run_engine(paths, cfg)
                st.session_state.last_run = datetime.now().strftime("%Y/%m/%d - %H:%M:%S")
                st.session_state.last_files_count = len(paths)


records = st.session_state.records


# ═════════════════════════════════════════════════════════════
# Empty State
# ═════════════════════════════════════════════════════════════

if not records:
    with card():
        st.subheader("هنوز گزارشی تولید نشده است")
        st.write(
            """
ابتدا فایل‌های Excel را بارگذاری کنید و دکمه «شروع پردازش» را بزنید.
بعد از پردازش، داشبورد، جدول نتایج، جزئیات شیت‌ها و خروجی Excel فعال می‌شود.
"""
        )

    st.stop()


# ═════════════════════════════════════════════════════════════
# محاسبات خلاصه
# ═════════════════════════════════════════════════════════════

total_records = len(records)
total_files = len(set(str(safe_get(r, COL["file"])) for r in records))

standard_counts = count_by(records, COL["standard_result"])
file_counts = count_by(records, COL["file_verdict"])

pass_count = standard_counts.get("قبول", 0)
fail_count = standard_counts.get("مردود", 0)
unknown_count = standard_counts.get("نامشخص", 0)

risks = [bounded_risk(safe_get(r, COL["risk"])) for r in records]
avg_risk = round(sum(risks) / len(risks), 1) if risks else 0
max_risk = max(risks) if risks else 0

if st.session_state.last_run:
    st.info(f"آخرین پردازش: {st.session_state.last_run}")


# ═════════════════════════════════════════════════════════════
# کارت‌های خلاصه
# ═════════════════════════════════════════════════════════════

m1, m2, m3, m4, m5 = st.columns(5)

with m1:
    show_metric_card(
        "فایل‌ها",
        total_files,
        f"{total_records} شیت بررسی شد",
    )

with m2:
    show_metric_card(
        "قبول استاندارد",
        pass_count,
        "بر اساس محاسبه مستقل",
    )

with m3:
    show_metric_card(
        "مردود استاندارد",
        fail_count,
        "دارای عدم انطباق اصلی",
    )

with m4:
    show_metric_card(
        "نامشخص",
        unknown_count,
        "نیازمند بررسی تکمیلی",
    )

with m5:
    show_metric_card(
        "میانگین ریسک",
        avg_risk,
        f"بیشترین ریسک: {max_risk}",
    )


st.divider()


# ═════════════════════════════════════════════════════════════
# Tabs
# ═════════════════════════════════════════════════════════════

tab_dashboard, tab_table, tab_detail, tab_report, tab_help = st.tabs(
    [
        "📊 داشبورد",
        "📋 جدول نتایج",
        "🔍 جزئیات شیت",
        "📤 خروجی گزارش",
        "📘 راهنما",
    ]
)


# ═════════════════════════════════════════════════════════════
# Tab 1: Dashboard
# ═════════════════════════════════════════════════════════════

with tab_dashboard:
    st.header("📊 داشبورد تحلیلی")

    c1, c2 = st.columns(2)

    with c1:
        show_count_progress(
            "نتیجه استاندارد",
            standard_counts,
        )

    with c2:
        show_count_progress(
            "اعتبارسنجی فایل",
            file_counts,
        )

    st.divider()

    c3, c4 = st.columns([1, 2])

    with c3:
        with card():
            st.subheader("خلاصه ریسک")
            show_risk(avg_risk)
            st.caption(f"بیشترین ریسک مشاهده‌شده: {max_risk}")

    with c4:
        show_count_chart(
            "نمودار نتیجه استاندارد",
            standard_counts,
        )

    st.divider()

    st.subheader("🔥 پرریسک‌ترین موارد")

    top_risk = sorted(
        records,
        key=lambda r: bounded_risk(safe_get(r, COL["risk"])),
        reverse=True,
    )[:10]

    top_columns = [
        COL["file"],
        COL["sheet"],
        COL["n"],
        COL["mtype"],
        COL["nominal"],
        COL["standard_result"],
        COL["file_verdict"],
        COL["form_result"],
        COL["risk"],
        COL["mean"],
        COL["sd"],
        COL["corrected_avg"],
        COL["t1_fail"],
        COL["t2_fail"],
    ]

    top_df = records_to_df(top_risk, top_columns)
    dataframe_with_risk(top_df)


# ═════════════════════════════════════════════════════════════
# Tab 2: Table
# ═════════════════════════════════════════════════════════════

with tab_table:
    st.header("📋 جدول کامل نتایج")

    base_df = records_to_df(records)

    f1, f2, f3, f4 = st.columns(4)

    files = sorted(set(str(safe_get(r, COL["file"])) for r in records))
    std_values = sorted(set(str(safe_get(r, COL["standard_result"])) for r in records))
    file_values = sorted(set(str(safe_get(r, COL["file_verdict"])) for r in records))
    mtypes = sorted(set(str(safe_get(r, COL["mtype"])) for r in records))

    with f1:
        selected_file = st.selectbox(
            "فایل",
            ["همه"] + files,
        )

    with f2:
        selected_std = st.selectbox(
            "نتیجه استاندارد",
            ["همه"] + std_values,
        )

    with f3:
        selected_file_verdict = st.selectbox(
            "اعتبارسنجی فایل",
            ["همه"] + file_values,
        )

    with f4:
        selected_mtype = st.selectbox(
            "نوع اندازه‌گیری",
            ["همه"] + mtypes,
        )

    risk_min, risk_max = st.slider(
        "بازه امتیاز ریسک",
        min_value=0,
        max_value=100,
        value=(0, 100),
        step=1,
    )

    search_text = st.text_input(
        "جستجو در نام فایل، شیت یا دلایل",
        placeholder="مثلاً نام محصول، استان، خطا، مشکوک...",
    )

    filtered = []

    for r in records:
        risk = bounded_risk(safe_get(r, COL["risk"]))

        if selected_file != "همه" and str(safe_get(r, COL["file"])) != selected_file:
            continue

        if selected_std != "همه" and str(safe_get(r, COL["standard_result"])) != selected_std:
            continue

        if selected_file_verdict != "همه" and str(safe_get(r, COL["file_verdict"])) != selected_file_verdict:
            continue

        if selected_mtype != "همه" and str(safe_get(r, COL["mtype"])) != selected_mtype:
            continue

        if not (risk_min <= risk <= risk_max):
            continue

        if search_text.strip():
            s = search_text.strip().lower()
            combined = " ".join(
                [
                    str(safe_get(r, COL["file"])),
                    str(safe_get(r, COL["sheet"])),
                    str(safe_get(r, COL["standard_reasons"])),
                    str(safe_get(r, COL["file_reasons"])),
                ]
            ).lower()

            if s not in combined:
                continue

        filtered.append(r)

    st.caption(f"تعداد رکوردهای قابل نمایش: {len(filtered)} از {len(records)}")

    preferred_columns = [
        COL["file"],
        COL["sheet"],
        COL["n"],
        COL["mtype"],
        COL["unit"],
        COL["nominal"],
        COL["standard_result"],
        COL["file_verdict"],
        COL["form_result"],
        COL["risk"],
        COL["mean"],
        COL["sd"],
        COL["corrected_avg"],
        COL["cv"],
        COL["t1_fail"],
        COL["t1_allowed"],
        COL["t2_fail"],
        COL["avg_req"],
        COL["t1_req"],
        COL["t2_req"],
        COL["net_calc_errors"],
        COL["shortage_errors"],
        COL["standard_reasons"],
        COL["file_reasons"],
    ]

    table_df = records_to_df(filtered, preferred_columns)
    dataframe_with_risk(table_df)

    with st.expander("نمایش همه ستون‌های خام"):
        raw_df = records_to_df(filtered)
        dataframe_with_risk(raw_df)


# ═════════════════════════════════════════════════════════════
# Tab 3: Detail
# ═════════════════════════════════════════════════════════════

with tab_detail:
    st.header("🔍 جزئیات فایل / شیت")

    options = []

    for idx, r in enumerate(records):
        file_name = safe_get(r, COL["file"])
        sheet_name = safe_get(r, COL["sheet"])
        std = safe_get(r, COL["standard_result"])
        fv = safe_get(r, COL["file_verdict"])
        risk = safe_get(r, COL["risk"])

        options.append(
            {
                "idx": idx,
                "label": f"{idx + 1} — {file_name} / {sheet_name} | استاندارد: {std} | فایل: {fv} | ریسک: {risk}",
            }
        )

    selected_label = st.selectbox(
        "رکورد موردنظر را انتخاب کنید",
        [x["label"] for x in options],
    )

    selected_idx = next(x["idx"] for x in options if x["label"] == selected_label)
    rec = records[selected_idx]

    st.divider()

    d1, d2, d3 = st.columns(3)

    with d1:
        with card():
            st.subheader("نتیجه استاندارد")
            show_status(safe_get(rec, COL["standard_result"]))
            st.caption("این نتیجه بر اساس محاسبه مستقل موتور است.")

    with d2:
        with card():
            st.subheader("اعتبارسنجی فایل")
            show_status(safe_get(rec, COL["file_verdict"]))
            st.caption("سلامت فایل، داده‌ها، فرمول‌ها و نشانه‌های غیرعادی بررسی می‌شود.")

    with d3:
        with card():
            st.subheader("امتیاز ریسک")
            show_risk(safe_get(rec, COL["risk"]))

    st.divider()

    st.subheader("مشخصات اصلی")

    main_keys = [
        COL["file"],
        COL["sheet"],
        COL["mtype"],
        COL["unit"],
        COL["nominal"],
        COL["n"],
        COL["form_result"],
        COL["risk"],
    ]

    for i in range(0, len(main_keys), 4):
        cols = st.columns(4)

        for col, key in zip(cols, main_keys[i:i + 4]):
            with col:
                show_detail_item(key, safe_get(rec, key))

    st.subheader("شاخص‌های استاندارد و آماری")

    standard_keys = [
        COL["tne"],
        COL["tne_pct"],
        COL["form_tolerance"],
        COL["mean"],
        COL["sd"],
        COL["corrected_avg"],
        COL["cv"],
        COL["avg_req"],
        COL["t1_req"],
        COL["t2_req"],
        COL["t1_fail"],
        COL["t1_allowed"],
        COL["t2_fail"],
    ]

    for i in range(0, len(standard_keys), 4):
        cols = st.columns(4)

        for col, key in zip(cols, standard_keys[i:i + 4]):
            with col:
                show_detail_item(key, safe_get(rec, key))

    st.subheader("شاخص‌های اعتبارسنجی فایل")

    validation_keys = [
        COL["net_calc_errors"],
        COL["shortage_errors"],
        COL["dup_rate"],
        COL["blocks"],
        COL["iqr"],
        COL["ex3"],
        COL["runs_p"],
        COL["round_pct"],
    ]

    for i in range(0, len(validation_keys), 4):
        cols = st.columns(4)

        for col, key in zip(cols, validation_keys[i:i + 4]):
            with col:
                show_detail_item(key, safe_get(rec, key))

    st.divider()

    st.subheader("دلایل و توضیحات")

    reason_col1, reason_col2 = st.columns(2)

    with reason_col1:
        std_reasons = split_reasons(safe_get(rec, COL["standard_reasons"]))
        show_reason_box("دلایل استاندارد", std_reasons)

    with reason_col2:
        file_reasons = split_reasons(safe_get(rec, COL["file_reasons"]))
        show_reason_box("دلایل اعتبارسنجی فایل", file_reasons)

    with st.expander("نمایش رکورد خام"):
        raw_one_df = records_to_df([rec])
        dataframe_with_risk(raw_one_df)


# ═════════════════════════════════════════════════════════════
# Tab 4: Report
# ═════════════════════════════════════════════════════════════

with tab_report:
    st.header("📤 خروجی گزارش Excel")

    with card():
        st.subheader("ساخت گزارش")
        st.write(
            """
گزارش خروجی با تابع `build_report` موتور محاسباتی ساخته می‌شود.
بنابراین ستون‌ها، منطق و رنگ‌بندی خروجی Excel مطابق موتور اصلی خواهد بود.
"""
        )

        st.metric("تعداد رکوردهای گزارش", len(records))

        try:
            excel_bytes = build_excel_bytes(records)

            report_name = f"inspection_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            st.download_button(
                label="⬇️ دانلود گزارش Excel",
                data=excel_bytes,
                file_name=report_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as ex:
            st.error("در ساخت گزارش Excel خطا رخ داد.")
            st.exception(ex)


# ═════════════════════════════════════════════════════════════
# Tab 5: Help
# ═════════════════════════════════════════════════════════════

with tab_help:
    st.header("📘 راهنمای تفسیر نتایج")

    with card():
        st.subheader("۱) نتیجه استاندارد")
        st.write(
            """
این نتیجه از محاسبات مستقل موتور به دست می‌آید و به نتیجه‌ای که داخل فرم Excel نوشته شده وابسته نیست.
معمولاً شامل الزام میانگین، آزمون T1 و آزمون T2 است.
"""
        )

    with card():
        st.subheader("۲) اعتبارسنجی فایل")
        st.write(
            """
این بخش کیفیت فایل را بررسی می‌کند؛ مانند خطا در محاسبات، نبود داده کافی، فرمول‌های ناقص،
تکرار غیرعادی داده‌ها، داده‌های پرت یا مغایرت بین ستون‌ها.
"""
        )

    with card():
        st.subheader("۳) نتیجه فرم")
        st.write(
            """
نتیجه‌ای است که در فایل Excel توسط کاربر یا فرم نوشته شده است.
این مقدار ممکن است با نتیجه استاندارد محاسبه‌شده توسط موتور متفاوت باشد.
"""
        )

    with card():
        st.subheader("۴) امتیاز ریسک")
        st.write(
            """
امتیاز ریسک عددی بین ۰ تا ۱۰۰ است.
هرچه این عدد بالاتر باشد، احتمال وجود خطا، مغایرت، داده‌سازی یا نیاز به بررسی کارشناسی بیشتر است.
"""
        )

    st.subheader("نحوه اجرا")

    st.code(
        """
pip install streamlit pandas openpyxl scipy
streamlit run app.py
        """.strip(),
        language="bash",
    )