import os
import base64
import pandas as pd
import streamlit as st
import plotly.express as px

# ---------------------------------------------------------
# إعدادات الصفحة الأساسية وتصميم الواجهة
# ---------------------------------------------------------
st.set_page_config(
    page_title="لوحة المؤشرات التنفيذية الحية",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
        html, body, [class*="css"] {
            font-family: 'Cairo', sans-serif;
            direction: rtl;
            text-align: right;
        }
        .main-header {
            font-size: 26px;
            font-weight: 900;
            color: #0f172a;
            border-bottom: 3px solid #2563eb;
            padding-bottom: 10px;
            margin-bottom: 25px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .badge {
            background-color: #dbeafe;
            color: #1d4ed8;
            padding: 6px 14px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 700;
        }
    </style>
""", unsafe_allow_html=True)

# رأس الصفحة التنفيذي
st.markdown("""
    <div class="main-header">
        <span>التقرير التحليلي التنفيذي للأداء (تفاعلي حي)</span>
        <span class="badge">معتمد للإدارة العليا</span>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# الشريط الجانبي للإعدادات ورفع الملفات
# ---------------------------------------------------------
st.sidebar.header("📁 إدارة البيانات والملفات")
uploaded_file = st.sidebar.file_uploader("ارفع ملف الأكسل أو CSV الخاص بك:", type=["xlsx", "xls", "csv"])

# ---------------------------------------------------------
# دوال معالجة وتحميل البيانات (ذكية وعالمية)
# ---------------------------------------------------------
def find_header_row(df_raw: pd.DataFrame) -> int:
    for idx, row in df_raw.iterrows():
        row_str = str(row.values).lower()
        looks_like_header = (
            "date" in row_str or "التاريخ" in str(row.values) or "day" in row_str or "month" in row_str or idx < 3
        )
        if looks_like_header and len(row.dropna()) >= 2:
            return idx
    return 0

def detect_date_column(df: pd.DataFrame) -> str:
    for col in df.columns:
        col_lower = str(col).lower()
        if "date" in col_lower or "التاريخ" in str(col) or "day" in col_lower or "month" in col_lower or "time" in col_lower:
            return col
    # إن لم يجد عمود تاريخ صريح، يأخذ أول عمود
    return df.columns[0]

@st.cache_data
def load_and_clean_data(file_source) -> pd.DataFrame:
    if isinstance(file_source, str):
        df_raw = pd.read_excel(file_source, header=None)
    else:
        if file_source.name.endswith('.csv'):
            df_raw = pd.read_csv(file_source, header=None)
        else:
            df_raw = pd.read_excel(file_source, header=None)
            
    header_row_idx = find_header_row(df_raw)
    
    if isinstance(file_source, str):
        df = pd.read_excel(file_source, header=header_row_idx)
    else:
        file_source.seek(0)
        if file_source.name.endswith('.csv'):
            df = pd.read_csv(file_source, header=header_row_idx)
        else:
            df = pd.read_excel(file_source, header=header_row_idx)

    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed")]

    date_col = detect_date_column(df)
    
    # محاولة تحويل عمود التاريخ لتاريخ صحيح، وإن فشل يحافظ على النصوص لتجنب الأخطاء
    try:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df = df.dropna(subset=[date_col]).sort_values(date_col).reset_index(drop=True)
    except:
        pass
        
    df.attrs["date_col"] = date_col
    return df

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, "data", "sales.xlsx")

try:
    if uploaded_file is not None:
        df = load_and_clean_data(uploaded_file)
        st.sidebar.success("تم تحميل وتصفية البيانات المرفوعة بنجاح!")
    elif os.path.exists(DEFAULT_EXCEL_PATH):
        df = load_and_clean_data(DEFAULT_EXCEL_PATH)
        st.sidebar.info("يتم عرض البيانات من الملف الافتراضي (sales.xlsx).")
    else:
        df = None
except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
    df = None

# ---------------------------------------------------------
# عرض المحتوى والتحليلات إذا توفرت البيانات
# ---------------------------------------------------------
if df is None:
    st.warning("الرجاء رفع ملف بيانات صالح من القائمة الجانبية أو توفير ملف sales.xlsx في مجلد data.")
else:
    date_col = df.attrs.get("date_col", df.columns[0])
    
    # تحويل كافة الأعمدة غير الرقمية (ما عدا عمود التاريخ) إلى أرقام لتجنب أي مشاكل
    for col in df.columns:
        if col != date_col:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if not numeric_cols:
        numeric_cols = [c for c in df.columns if c != date_col]

    st.sidebar.header("🎛️ خيارات التحليل المتقدم")
    selected_metric = st.sidebar.selectbox("اختر المؤشر الرئيسي للتحليل:", numeric_cols, index=0)
    
    compare_option = st.sidebar.checkbox("إضافة مؤشر ثانٍ للمقارنة البيانية")
    selected_metric_2 = None
    if compare_option:
        remaining_cols = [c for c in numeric_cols if c != selected_metric]
        if remaining_cols:
            selected_metric_2 = st.sidebar.selectbox("اختر المؤشر الثاني للمقارنة:", remaining_cols)

    total_val = df[selected_metric].sum()
    avg_val = df[selected_metric].mean()
    max_val = df[selected_metric].max()
    min_val = df[selected_metric].min()

    # استخراج أفضل تاريخ أو عنصر سُميت به القيمة العظمى ذكياً
    try:
        max_row = df.loc[df[selected_metric].idxmax()]
        best_time_label = str(max_row[date_col]).split()[0]
    except:
        best_time_label = "فترة الذروة"

    # 1) التحليل البصري (Plotly)
    st.subheader("أولاً: التحليل البصري التفاعلي للأداء")
    
    fig = px.line(
        df, x=date_col, y=selected_metric,
        markers=True,
        labels={selected_metric: "القيمة / المؤشر", date_col: "النطاق الزمني"},
        title=f"حركة الأداء لـ: {selected_metric}"
    )
    fig.update_traces(line=dict(width=3, color="#2563eb"), marker=dict(size=6, color="#2563eb"))
    
    if selected_metric_2:
        fig.add_scatter(
            x=df[date_col], y=df[selected_metric_2],
            mode='lines+markers', name=selected_metric_2,
            line=dict(width=3, color="#059669"),
            marker=dict(size=6, color="#059669")
        )

    fig.update_layout(
        plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
        font=dict(family="Cairo", size=12), hovermode="x unified",
        margin=dict(t=40, b=20, l=20, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig, use_container_width=True)

    # 2) لوحة المؤشرات الإحصائية
    st.subheader("ثانياً: لوحة المؤشرات الإحصائية الرئيسية")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric(label=f"إجمالي المجموع ({selected_metric})", value=f"{total_val:,.2f}")
    with c2:
        st.metric(label="المتوسط الحسابي", value=f"{avg_val:,.2f}")
    with c3:
        st.metric(label="القيمة العليا (الذروة)", value=f"{max_val:,.2f}")
    with c4:
        st.metric(label="القيمة الدنيا", value=f"{min_val:,.2f}")

    # 3) القراءة التحليلية الذكية والتلقائية
    st.subheader("ثالثاً: القراءة التحليلية والتوصيات التنفيذية")
    analysis_text = f"""
    أظهرت البيانات المحدثة استقراراً وتحليلاً تفاعلياً لـ <strong>{selected_metric}</strong>، حيث بلغ إجمالي الحجم الكلي <strong>{total_val:,.2f}</strong> بمتوسط حسابي يستقر عند <strong>{avg_val:,.2f}</strong>.<br>
    • سجل المؤشر <strong>أعلى أداء (ذروة)</strong> بواقع <strong>{max_val:,.2f}</strong> بتاريخ أو نقطة قياس <strong>{best_time_label}</strong>.<br>
    • بلغ أدنى مستوى مسجل نحو <strong>{min_val:,.2f}</strong>.<br><br>
    <strong>التوصية الإدارية:</strong> يوصى بدراسة العوامل التي ساهمت في تحقيق الذروة بتاريخ {best_time_label} لتعميم التجارب الناجحة، مع متابعة الفترات المنخفضة لتحسين كفاءة التشغيل.
    """
    
    st.markdown(f"""
    <div style="background-color: #f8fafc; border-right: 5px solid #059669; padding: 20px 25px; border-radius: 0 12px 12px 0; line-height: 1.9; font-size: 15px; color: #334155; border: 1px solid #e2e8f0; box-shadow: 0 2px 8px rgba(0,0,0,0.01);">
        <strong>قراءة تنفيذية ذكية:</strong> {analysis_text}
    </div>
    """, unsafe_allow_html=True)

    # ---------------------------------------------------------
    # 4) ميزة تصدير التقرير التنفيذي (جاهز للطباعة / PDF)
    # ---------------------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.header("🖨️ التصدير والطباعة الرسمية")
    
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
            .analysis {{ background: #f8fafc; border-right: 4px solid #059669; padding: 15px; border-radius: 0 8px 8px 0; line-height: 1.8; font-size: 13px; color: #334155; border: 1px solid #e2e8f0; }}
            .footer {{ margin-top: 35px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 15px; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>التقرير التحليلي التنفيذي للأداء ({selected_metric})</h1>
            <div class="badge">معتمد للإدارة العليا</div>
        </div>

        <div class="section-title">أولاً: لوحة المؤشرات الإحصائية الرئيسية</div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="title">إجمالي المجموع</div>
                <div class="value">{total_val:,.2f}</div>
            </div>
            <div class="stat-card">
                <div class="title">المتوسط الحسابي</div>
                <div class="value">{avg_val:,.2f}</div>
            </div>
            <div class="stat-card">
                <div class="title">القيمة العليا (الذروة)</div>
                <div class="value">{max_val:,.2f}</div>
            </div>
            <div class="stat-card">
                <div class="title">القيمة الدنيا</div>
                <div class="value">{min_val:,.2f}</div>
            </div>
        </div>

        <div class="section-title">ثانياً: القراءة التحليلية والتوصيات التنفيذية</div>
        <div class="analysis">
            {analysis_text}
        </div>

        <div class="footer">
            نظام التحليل المؤسسي والتقارير التنفيذية الحية • تم الإصدار رسمياً
        </div>
    </body>
    </html>
    """

    st.sidebar.download_button(
        label="📥 تحميل التقرير كملف HTML معتمد",
        data=html_export_content,
        file_name="Executive_Report.html",
        mime="text/html",
        help="اضغط لتحميل التقرير، ويمكنك فتحه في أي متصفح وحفظه مباشرة بصيغة PDF بضغط Ctrl + P"
    )

    # تذييل الصفحة
    st.markdown("<br><hr><center><span style='color: #94a3b8; font-size: 13px; font-weight: 600;'>نظام التحليل المؤسسي والتقارير التنفيذية الحية • تم الإصدار بنجاح</span></center>", unsafe_allow_html=True)
