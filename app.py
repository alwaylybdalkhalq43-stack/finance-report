"""
=========================================================
لوحة المؤشرات التنفيذية الحية - Enterprise Edition
تحليل ذكي لأي ملف بيانات (Excel / CSV) مع رسوم بيانية
تفاعلية وقراءة تحليلية تلقائية تُبنى من البيانات نفسها.
=========================================================
"""

import base64
from typing import Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

# ---------------------------------------------------------
# إعدادات الصفحة والهوية البصرية
# ---------------------------------------------------------

st.set_page_config(
    page_title="لوحة المؤشرات التنفيذية الحية",
    page_icon="📊",
    layout="wide",
)

COLOR_PRIMARY = "#2563eb"
COLOR_SECONDARY = "#059669"
COLOR_DANGER = "#dc2626"
COLOR_TEXT_DARK = "#0f172a"
COLOR_TEXT_MUTED = "#64748b"

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
        html, body, [class*="css"] {
            font-family: 'Cairo', sans-serif;
            direction: rtl;
            text-align: right;
        }
        .main-header {
            font-size: 26px; font-weight: 900; color: #0f172a;
            border-bottom: 3px solid #2563eb; padding-bottom: 10px;
            margin-bottom: 25px; display: flex;
            justify-content: space-between; align-items: center;
        }
        .badge {
            background-color: #dbeafe; color: #1d4ed8;
            padding: 6px 14px; border-radius: 8px;
            font-size: 13px; font-weight: 700;
        }
        .insight-box {
            background-color: #f8fafc; border-right: 5px solid #059669;
            padding: 20px 25px; border-radius: 0 12px 12px 0;
            line-height: 2; font-size: 15px; color: #334155;
            border: 1px solid #e2e8f0; box-shadow: 0 2px 8px rgba(0,0,0,0.01);
        }
        div[data-testid="stMetric"] {
            background-color: #f8fafc; border: 1px solid #e2e8f0;
            border-radius: 10px; padding: 12px 16px;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <span>التقرير التحليلي التنفيذي للأداء (تفاعلي حي)</span>
        <span class="badge">معتمد للإدارة العليا</span>
    </div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# 1) تحميل وتنظيف البيانات (تدعم أي ملف Excel/CSV بدون افتراض شكل ثابت)
# ---------------------------------------------------------

def _find_header_row(df_raw: pd.DataFrame) -> int:
    """يكتشف صف العناوين الحقيقي في حال وجود صفوف شعار/عنوان قبله."""
    for idx, row in df_raw.iterrows():
        row_str = str(row.values).lower()
        looks_like_header = (
            "date" in row_str
            or "التاريخ" in str(row.values)
            or "day" in row_str
            or idx < 3
        )
        if looks_like_header and len(row.dropna()) >= 2:
            return idx
    return 0


def _detect_date_column(df: pd.DataFrame) -> Optional[str]:
    """
    يحاول اكتشاف عمود تاريخ فعلي. يرجع None إن لم يوجد أي عمود يشبه تاريخاً،
    بدل افتراض أن العمود الأول هو التاريخ دائماً (كما كانت النسخة السابقة تفعل).
    """
    for col in df.columns:
        col_lower = str(col).lower()
        if "date" in col_lower or "التاريخ" in str(col) or "day" in col_lower or "تاريخ" in str(col):
            parsed = pd.to_datetime(df[col], errors="coerce")
            if parsed.notna().mean() >= 0.7:  # 70%+ من القيم قابلة للتحويل فعلاً
                return col

    # محاولة أخيرة: أول عمود قابل للتحويل لتاريخ بنسبة عالية
    for col in df.columns[:2]:
        parsed = pd.to_datetime(df[col], errors="coerce")
        if parsed.notna().mean() >= 0.7:
            return col

    return None


@st.cache_data(show_spinner=False)
def load_and_clean_data(file_source) -> tuple[pd.DataFrame, Optional[str]]:
    """
    يقرأ أي ملف Excel/CSV، يكتشف صف العناوين وعمود التاريخ تلقائياً،
    وينظف الصفوف/الأعمدة الفارغة. يرجع (البيانات النظيفة, اسم عمود التاريخ أو None).
    """
    is_path = isinstance(file_source, str)
    is_csv = (file_source if is_path else file_source.name).lower().endswith(".csv")

    if not is_path:
        file_source.seek(0)

    if is_csv:
        df_raw = pd.read_csv(file_source, header=None, encoding_errors="ignore")
    else:
        df_raw = pd.read_excel(file_source, header=None)

    header_row_idx = _find_header_row(df_raw)

    if not is_path:
        file_source.seek(0)

    if is_csv:
        df = pd.read_csv(file_source, header=header_row_idx, encoding_errors="ignore")
    else:
        df = pd.read_excel(file_source, header=header_row_idx)

    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed")]
    df.columns = [str(c).strip() for c in df.columns]

    date_col = _detect_date_column(df)
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df = df.dropna(subset=[date_col]).sort_values(date_col).reset_index(drop=True)

    return df, date_col


# ---------------------------------------------------------
# 2) دوال التنسيق والتحليل الذكي (قلب الميزة الجديدة)
# ---------------------------------------------------------

def format_compact_number(value: float) -> str:
    """6,500,000 -> 6.5M | 850,000 -> 850K | غير ذلك يعرض بفواصل عادية."""
    if pd.isna(value):
        return "—"
    abs_value = abs(value)
    if abs_value >= 1_000_000:
        return f"{value / 1_000_000:,.2f}M"
    if abs_value >= 1_000:
        return f"{value / 1_000:,.1f}K"
    return f"{value:,.2f}"


def build_insight(df: pd.DataFrame, date_col: Optional[str], metric: str) -> dict:
    """
    يستخرج قراءة تحليلية كاملة من البيانات: الاتجاه العام، نسبة التغير،
    التذبذب، وأهم نقطة عليا/دنيا - كل ده محسوب فعلياً من البيانات
    (مش نص ثابت مكرر زي النسخة القديمة).
    """
    series = df[metric].dropna()

    total_val = series.sum()
    avg_val = series.mean()
    max_val = series.max()
    min_val = series.min()
    std_val = series.std() if len(series) > 1 else 0.0
    volatility_ratio = (std_val / avg_val * 100) if avg_val else 0.0

    max_idx = series.idxmax()
    min_idx = series.idxmin()
    max_label = df.loc[max_idx, date_col].strftime("%Y-%m-%d") if date_col else f"صف {max_idx}"
    min_label = df.loc[min_idx, date_col].strftime("%Y-%m-%d") if date_col else f"صف {min_idx}"

    # الاتجاه العام: مقارنة نصف البيانات الأول بالنصف الثاني (أكثر ثباتاً من أول/آخر نقطة فقط)
    half = max(len(series) // 2, 1)
    first_half_avg = series.iloc[:half].mean()
    second_half_avg = series.iloc[half:].mean()
    pct_change = (
        ((second_half_avg - first_half_avg) / first_half_avg * 100)
        if first_half_avg else 0.0
    )

    if pct_change > 5:
        trend_label, trend_word = "تصاعدي", "ارتفاعاً"
    elif pct_change < -5:
        trend_label, trend_word = "تنازلي", "انخفاضاً"
    else:
        trend_label, trend_word = "مستقر", "استقراراً نسبياً"

    if volatility_ratio > 40:
        volatility_label = "تذبذب مرتفع يستدعي المتابعة الدقيقة"
    elif volatility_ratio > 20:
        volatility_label = "تذبذب متوسط ضمن الحدود المقبولة"
    else:
        volatility_label = "استقرار جيد في التوزيع العام للقيم"

    return {
        "total": total_val, "avg": avg_val, "max": max_val, "min": min_val,
        "std": std_val, "volatility_ratio": volatility_ratio,
        "max_label": max_label, "min_label": min_label,
        "pct_change": pct_change, "trend_label": trend_label,
        "trend_word": trend_word, "volatility_label": volatility_label,
    }


def build_narrative(metric: str, insight: dict, metric2: Optional[str] = None, insight2: Optional[dict] = None) -> str:
    """يبني نص القراءة التحليلية التنفيذية ديناميكياً من الأرقام الفعلية."""

    arrow = "▲" if insight["pct_change"] >= 0 else "▼"

    text = (
        f"سجّل مؤشر <strong>{metric}</strong> {insight['trend_word']} خلال الفترة المشمولة بالتحليل، "
        f"بنسبة تغير تقارب <strong>{arrow} {abs(insight['pct_change']):,.1f}%</strong> "
        f"بين النصف الأول والثاني من البيانات (اتجاه عام: <strong>{insight['trend_label']}</strong>).<br><br>"
        f"بلغت أعلى قيمة مسجلة <strong>{format_compact_number(insight['max'])}</strong> "
        f"بتاريخ <strong>{insight['max_label']}</strong>، بينما سجلت أدنى قيمة "
        f"<strong>{format_compact_number(insight['min'])}</strong> بتاريخ <strong>{insight['min_label']}</strong>، "
        f"واستقر المتوسط العام عند <strong>{format_compact_number(insight['avg'])}</strong>.<br><br>"
        f"<strong>تقييم الاستقرار:</strong> {insight['volatility_label']} "
        f"(نسبة التذبذب حوالي {insight['volatility_ratio']:,.1f}% من المتوسط)."
    )

    if metric2 and insight2:
        ratio = (insight["avg"] / insight2["avg"]) if insight2["avg"] else 0
        text += (
            f"<br><br><strong>قراءة مقارنة:</strong> بمقارنة <strong>{metric}</strong> بمؤشر "
            f"<strong>{metric2}</strong>، فإن متوسط الأول يعادل تقريباً "
            f"<strong>{ratio:,.1f}x</strong> من متوسط الثاني، "
            f"مع اتجاه {insight2['trend_word']} للمؤشر الثاني خلال نفس الفترة."
        )

    text += (
        "<br><br><strong>التوصية الإدارية:</strong> "
        + (
            "يُوصى بمتابعة يومية دقيقة نظراً لارتفاع نسبة التذبذب وضمان استجابة سريعة لأي انحراف عن المسار المستهدف."
            if insight["volatility_ratio"] > 40 else
            "يُوصى بالاستمرار في وتيرة المتابعة الحالية مع مراجعة دورية للمؤشر لضمان استدامة الأداء."
        )
    )

    return text


# ---------------------------------------------------------
# 3) بناء الرسم البياني الاحترافي (Plotly)
# ---------------------------------------------------------

def build_chart(df: pd.DataFrame, x_col: str, metric: str, insight: dict,
                 metric2: Optional[str] = None, is_time_series: bool = True) -> go.Figure:
    """يبني رسماً بيانياً تفاعلياً بمستوى تنفيذي: تسميات القمة/القاع، تنسيق مختصر للأرقام، Range Slider."""

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df[x_col], y=df[metric], name=metric, mode="lines+markers",
        line=dict(width=3, color=COLOR_PRIMARY, shape="spline", smoothing=0.3),
        marker=dict(size=6, color=COLOR_PRIMARY),
        fill="tozeroy", fillcolor="rgba(37, 99, 235, 0.07)",
    ))

    if metric2:
        fig.add_trace(go.Scatter(
            x=df[x_col], y=df[metric2], name=metric2, mode="lines+markers",
            line=dict(width=3, color=COLOR_SECONDARY, shape="spline", smoothing=0.3),
            marker=dict(size=6, color=COLOR_SECONDARY),
            fill="tozeroy", fillcolor="rgba(5, 150, 105, 0.06)",
        ))

    # تسميات القمة والقاع
    fig.add_annotation(
        x=df.loc[df[metric].idxmax(), x_col], y=insight["max"],
        text=f"أعلى: {format_compact_number(insight['max'])}",
        showarrow=True, arrowhead=2, ax=0, ay=-35,
        bgcolor=COLOR_PRIMARY, font=dict(color="white", size=11), bordercolor=COLOR_PRIMARY,
    )
    fig.add_annotation(
        x=df.loc[df[metric].idxmin(), x_col], y=insight["min"],
        text=f"أدنى: {format_compact_number(insight['min'])}",
        showarrow=True, arrowhead=2, ax=0, ay=35,
        bgcolor=COLOR_TEXT_MUTED, font=dict(color="white", size=11), bordercolor=COLOR_TEXT_MUTED,
    )

    fig.update_layout(
        title=f"حركة الأداء الزمني لـ: {metric}" if is_time_series else f"توزيع القيم حسب {x_col}",
        plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
        font=dict(family="Cairo", size=12, color=COLOR_TEXT_DARK),
        hovermode="x unified",
        margin=dict(t=60, b=20, l=20, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
        yaxis=dict(tickformat="~s", gridcolor="#eef1f5", zeroline=False),
        xaxis=dict(showgrid=False),
    )

    if is_time_series:
        fig.update_xaxes(rangeslider_visible=True, rangeslider_thickness=0.06)

    return fig


# ---------------------------------------------------------
# 4) الشريط الجانبي: رفع الملفات
# ---------------------------------------------------------

st.sidebar.header("📁 إدارة البيانات والملفات")
uploaded_file = st.sidebar.file_uploader(
    "ارفع ملف الإكسل أو CSV الخاص بك:", type=["xlsx", "xls", "csv"]
)

import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, "data", "sales.xlsx")

df: Optional[pd.DataFrame] = None
date_col: Optional[str] = None

try:
    if uploaded_file is not None:
        df, date_col = load_and_clean_data(uploaded_file)
        st.sidebar.success(f"تم تحميل {len(df):,} صف بنجاح.")
    elif os.path.exists(DEFAULT_EXCEL_PATH):
        df, date_col = load_and_clean_data(DEFAULT_EXCEL_PATH)
        st.sidebar.info("يتم عرض البيانات من الملف الافتراضي (sales.xlsx).")
except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
    df = None


# ---------------------------------------------------------
# 5) عرض التحليل الكامل
# ---------------------------------------------------------

if df is None:
    st.warning("الرجاء رفع ملف بيانات صالح (Excel أو CSV) من القائمة الجانبية للبدء في التحليل.")
elif df.empty:
    st.error("الملف تم تحميله لكنه لا يحتوي على أي بيانات صالحة بعد التنظيف.")
else:
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if not numeric_cols:
        for col in df.columns:
            if col != date_col:
                converted = pd.to_numeric(df[col], errors="coerce")
                if converted.notna().mean() >= 0.5:
                    df[col] = converted.fillna(0)
        numeric_cols = df.select_dtypes(include="number").columns.tolist()

    if not numeric_cols:
        st.error("لم يتم العثور على أي أعمدة رقمية صالحة للتحليل في هذا الملف.")
        st.dataframe(df.head(20), use_container_width=True)
        st.stop()

    is_time_series = date_col is not None
    x_col = date_col if is_time_series else (
        df.select_dtypes(exclude="number").columns[0] if len(df.select_dtypes(exclude="number").columns) else df.index.name or "index"
    )
    if not is_time_series and x_col == "index":
        df = df.reset_index()

    if not is_time_series:
        st.info("لم يتم اكتشاف عمود تاريخ في البيانات، سيتم عرض المؤشرات مرتبة حسب ترتيب الصفوف.")

    # --- خيارات التحليل ---
    st.sidebar.header("🎛️ خيارات التحليل المتقدم")
    selected_metric = st.sidebar.selectbox("اختر المؤشر الرئيسي للتحليل:", numeric_cols, index=0)

    compare_option = st.sidebar.checkbox("إضافة مؤشر ثانٍ للمقارنة البيانية")
    selected_metric_2 = None
    if compare_option:
        remaining_cols = [c for c in numeric_cols if c != selected_metric]
        if remaining_cols:
            selected_metric_2 = st.sidebar.selectbox("اختر المؤشر الثاني للمقارنة:", remaining_cols)

    insight = build_insight(df, date_col, selected_metric)
    insight2 = build_insight(df, date_col, selected_metric_2) if selected_metric_2 else None

    # --- 1) الرسم البياني ---
    st.subheader("أولاً: التحليل البصري التفاعلي للأداء")
    fig = build_chart(df, x_col, selected_metric, insight, selected_metric_2, is_time_series)
    st.plotly_chart(fig, use_container_width=True)

    # --- 2) لوحة المؤشرات مع نسبة التغير (Delta) ---
    st.subheader("ثانياً: لوحة المؤشرات الإحصائية الرئيسية")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"إجمالي المجموع ({selected_metric})", format_compact_number(insight["total"]))
    c2.metric("المتوسط الحسابي", format_compact_number(insight["avg"]),
              delta=f"{insight['pct_change']:+.1f}% مقارنة بالنصف الأول", delta_color="normal")
    c3.metric("القيمة العظمى", format_compact_number(insight["max"]))
    c4.metric("القيمة الصغرى", format_compact_number(insight["min"]))

    # --- 3) القراءة التحليلية التلقائية ---
    st.subheader("ثالثاً: القراءة التحليلية والتوصيات التنفيذية")
    narrative = build_narrative(selected_metric, insight, selected_metric_2, insight2)
    st.markdown(f"""<div class="insight-box"><strong>قراءة تنفيذية متقدمة:</strong><br>{narrative}</div>""",
                unsafe_allow_html=True)

    # --- 4) استكشاف البيانات الخام (اختياري) ---
    with st.expander("📋 عرض البيانات الخام بعد التنظيف"):
        st.dataframe(df, use_container_width=True)

    # ---------------------------------------------------------
    # 5) تصدير التقرير كملف HTML تفاعلي (يتضمن الرسم البياني فعلياً)
    # ---------------------------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.header("🖨️ التصدير والطباعة الرسمية")

    chart_html = pio.to_html(fig, include_plotlyjs="cdn", full_html=False, config={"displaylogo": False})

    html_export_content = f"""
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>التقرير التنفيذي المعتمد - {selected_metric}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
        body {{ font-family: 'Cairo', sans-serif; background-color: #ffffff; padding: 30px; color: #1e293b; direction: rtl; text-align: right; }}
        .header {{ border-bottom: 3px solid #2563eb; padding-bottom: 15px; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }}
        .header h1 {{ font-size: 22px; color: #0f172a; margin: 0; }}
        .badge {{ background-color: #dbeafe; color: #1d4ed8; padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: bold; }}
        .section-title {{ font-size: 15px; color: #0f172a; margin-top: 25px; margin-bottom: 12px; font-weight: bold; border-right: 4px solid #2563eb; padding-right: 10px; }}
        .stats-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }}
        .stat-card {{ background: #f8fafc; border: 1px solid #cbd5e1; padding: 15px; border-radius: 8px; text-align: center; }}
        .stat-card .title {{ font-size: 11px; color: #64748b; margin-bottom: 5px; }}
        .stat-card .value {{ font-size: 16px; font-weight: bold; color: #0f172a; }}
        .analysis {{ background: #f8fafc; border-right: 4px solid #059669; padding: 15px; border-radius: 0 8px 8px 0; line-height: 1.9; font-size: 13px; color: #334155; border: 1px solid #e2e8f0; }}
        .footer {{ margin-top: 35px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 15px; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>التقرير التحليلي التنفيذي للأداء ({selected_metric})</h1>
        <div class="badge">معتمد للإدارة العليا</div>
    </div>

    <div class="section-title">أولاً: التحليل البصري التفاعلي</div>
    {chart_html}

    <div class="section-title">ثانياً: لوحة المؤشرات الإحصائية الرئيسية</div>
    <div class="stats-grid">
        <div class="stat-card"><div class="title">إجمالي المجموع</div><div class="value">{format_compact_number(insight['total'])}</div></div>
        <div class="stat-card"><div class="title">المتوسط الحسابي</div><div class="value">{format_compact_number(insight['avg'])}</div></div>
        <div class="stat-card"><div class="title">القيمة العظمى</div><div class="value">{format_compact_number(insight['max'])}</div></div>
        <div class="stat-card"><div class="title">القيمة الصغرى</div><div class="value">{format_compact_number(insight['min'])}</div></div>
    </div>

    <div class="section-title">ثالثاً: القراءة التحليلية والتوصيات التنفيذية</div>
    <div class="analysis">{narrative}</div>

    <div class="footer">نظام التحليل المؤسسي والتقارير التنفيذية الحية • تم الإصدار رسمياً</div>
</body>
</html>
"""

    st.sidebar.download_button(
        label="📥 تحميل التقرير كملف HTML معتمد (يتضمن الرسم التفاعلي)",
        data=html_export_content,
        file_name="Executive_Report.html",
        mime="text/html",
        help="التقرير يحتوي على الرسم البياني التفاعلي نفسه، ويمكن حفظه كـ PDF عبر Ctrl+P من المتصفح.",
    )

    st.markdown(
        "<br><hr><center><span style='color: #94a3b8; font-size: 13px; font-weight: 600;'>"
        "نظام التحليل المؤسسي والتقارير التنفيذية الحية • تم الإصدار بنجاح</span></center>",
        unsafe_allow_html=True,
    )
