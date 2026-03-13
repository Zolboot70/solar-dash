import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Custom CSS
st.markdown(
    """
    <style>
        [data-testid="stMetricValue"] { font-size: 1.4rem !important; line-height: 1.2 !important; }
        [data-testid="stMetricLabel"] { font-size: 0.9rem !important; }
    </style>
    """,
    unsafe_allow_html=True
)

st.set_page_config(page_title="Нарны цахилгаан станцын үйлдвэрлэл", layout="wide")

st.title("Нарны цахилгаан станцын үйлдвэрлэл")
st.markdown("PVMS-ын стандарт хэлбэртэй файл оруулна уу")

uploaded_file = st.file_uploader(
    "НЦС-ын үйлдвэрлэлийн мэдээг оруулна уу",
    type=["xlsx"],
    help="Sheet1 байх ёстой"
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", skiprows=1)

        # Aggressive column cleaning
        df.columns = (
            df.columns.astype(str)
            .str.strip()
            .str.replace(r'\s+', ' ', regex=True)
            .str.replace(r'[^a-zA-Z0-9 ()°%°/.,-]', '', regex=True)
        )

        # Debug columns (remove later if you want)
        # st.caption("Detected columns: " + ", ".join(df.columns.tolist()))

        # ────────────────────────────────────────
        # DATE HANDLING - smart type detection
        # ────────────────────────────────────────
        date_col = None
        for col in df.columns:
            col_lower = col.lower().strip()
            if "statistical" in col_lower or "period" in col_lower or "date" in col_lower or "огноо" in col_lower:
                date_col = col
                break

        if date_col is None:
            st.error("Огнооны багана олдсонгүй! ('Statistical Period', 'Date', 'Огноо' гэх мэт байх ёстой)")
            st.stop()

        # Decide conversion based on content
        sample = df[date_col].head(5).astype(str).str.strip()
        is_serial = sample.str.match(r'^\d{5,6}$').all() or sample.str.isdigit().mean() > 0.8

        if is_serial:
            df["Date"] = pd.to_datetime(df[date_col], unit='D', origin='1899-12-30', errors='coerce')
        else:
            df["Date"] = pd.to_datetime(df[date_col], errors='coerce')

        df = df.drop(columns=[date_col], errors='ignore')
        df = df.dropna(subset=["Date"])
        df = df.sort_values("Date")

        # ────────────────────────────────────────
        # Renaming other columns
        # ────────────────────────────────────────
        rename_keywords = {
            "theoretical|боломжит|yield": "Theoretical_Yield",
            "inverter": "Inverter_Yield",
            "pv yield|pv": "PV_Yield",
            "peak|оргил|max": "Peak_Power",
            "co2|co₂|carbon": "CO2_Avoided",
            "charge|цэнэглэсэн": "Charge",
            "discharge|нийлүүлсэн": "Discharge",
        }

        rename_map = {}
        for col in df.columns:
            col_lower = col.lower()
            for pattern, new_name in rename_keywords.items():
                if any(kw in col_lower for kw in pattern.split('|')):
                    rename_map[col] = new_name
                    break

        df = df.rename(columns=rename_map)

        # ────────────────────────────────────────
        # Sidebar date filter
        # ────────────────────────────────────────
        st.sidebar.header("МАК НЦС")

        min_date = df["Date"].min().date()
        max_date = df["Date"].max().date()

        date_range = st.sidebar.date_input(
            "Хугацаа",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

        if len(date_range) == 2:
            start_date, end_date = date_range
            mask = (df["Date"].dt.date >= start_date) & (df["Date"].dt.date <= end_date)
            df_filtered = df.loc[mask].copy()
        else:
            df_filtered = df.copy()

        # ────────────────────────────────────────
        # Totals (safe)
        # ────────────────────────────────────────
        total_theo   = df_filtered.get("Theoretical_Yield",   pd.Series(0)).sum()
        total_inv    = df_filtered.get("Inverter_Yield",      pd.Series(0)).sum()
        total_co2    = df_filtered.get("CO2_Avoided",         pd.Series(0)).sum()
        total_charge = df_filtered.get("Charge",              pd.Series(0)).sum()
        total_disch  = df_filtered.get("Discharge",           pd.Series(0)).sum()
        avg_peak     = df_filtered.get("Peak_Power",          pd.Series(0)).mean() or 0.0

        # ────────────────────────────────────────
        # Grid input
        # ────────────────────────────────────────
        st.subheader("Сүлжээнээс нийлүүлсэн эрчим хүч")
        grid_energy = st.number_input(
            "Сүлжээнээс нийлүүлсэн эрчим хүчийг оруулна уу (кВт·ц)",
            min_value=0.0, value=0.0, step=100.0, format="%.0f"
        )

        # ────────────────────────────────────────
        # Metrics
        # ────────────────────────────────────────
        cols = st.columns(6)
        total_consumption = grid_energy + total_inv - total_charge + total_disch - 1200

        cols[0].metric("Үйлдвэрлэх боломжит эрчим хүч", f"{total_theo:,.0f} кВт·ц")
        cols[1].metric("Үйлдвэрлэсэн эрчим хүч",       f"{total_inv:,.0f} кВт·ц")
        cols[2].metric("CO₂ бууруулсан",                f"{total_co2:,.2f} тн")
        cols[3].metric("Нийт ЦЭХ хэрэглээ",            f"{total_consumption:,.0f} кВт·ц")
        cols[4].metric("Батарейнаас нийлүүлсэн",        f"{total_disch:,.0f} кВт·ц")
        cols[5].metric("Max чадлын дундаж",             f"{avg_peak:,.1f} кВт")

        st.markdown("---")

        # Tabs & charts (same as before - omitted for brevity, copy from your previous version)

        # ... (paste your tab1 to tab5 and data table code here - no changes needed there)

    except Exception as e:
        st.error(f"Файлыг уншихад алдаа гарлаа:\n{str(e)}")
        st.info("Шалгах зүйлс:\n• Sheet1 байгаа эсэх\n• 'Statistical Period' багана байгаа эсэх\n• Огноо формат зөв эсэх")

else:
    st.info("Файл сонгоно уу")

st.markdown("---")
st.caption("МАК НЦС үйлдвэрлэл")
