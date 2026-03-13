import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ────────────────────────────────────────
# Custom CSS - smaller metrics
# ────────────────────────────────────────
st.markdown(
    """
    <style>
        [data-testid="stMetricValue"] {
            font-size: 1.4rem !important;
            line-height: 1.2 !important;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.9rem !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.set_page_config(
    page_title="Нарны цахилгаан станцын үйлдвэрлэл",
    layout="wide"
)

st.title("Нарны цахилгаан станцын үйлдвэрлэл")
st.markdown("PVMS-ын стандарт хэлбэртэй файл оруулна уу")

uploaded_file = st.file_uploader(
    "НЦС-ын үйлдвэрлэлийн мэдээг оруулна уу",
    type=["xlsx"],
    help="Файл дотор 'Sheet1' байх ёстой"
)

if uploaded_file is not None:
    try:
        # ────────────────────────────────────────
        # READ DATA
        # ────────────────────────────────────────
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", skiprows=1)

        # Clean column names (remove extra spaces)
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # Flexible renaming keywords
        rename_keywords = {
            "date|period|statistical|огноо|хугацаа|report date": "Date",
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
        # DATE COLUMN HANDLING - more robust detection
        # ────────────────────────────────────────
        possible_date_keywords = [
            "date", "period", "statistical", "statistical period",
            "огноо", "хугацаа", "report date", "time"
        ]

        date_col = None
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if any(kw in col_lower for kw in possible_date_keywords):
                date_col = col
                break

        if date_col:
            df["Date"] = pd.to_datetime(df[date_col], errors="coerce")
            df = df.drop(columns=[date_col], errors="ignore")
            # st.info(f"Огнооны багана: {date_col} → Date болголоо")
        else:
            st.error("Огнооны багана олдсонгүй! ('Statistical Period', 'Date', 'Огноо' гэх мэт байх ёстой)")
            st.stop()

        df = df.dropna(subset=["Date"])
        df = df.sort_values("Date")

        # ────────────────────────────────────────
        # SIDEBAR FILTER
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
        # SAFE TOTALS
        # ────────────────────────────────────────
        total_theo   = df_filtered.get("Theoretical_Yield",   pd.Series(0)).sum()
        total_inv    = df_filtered.get("Inverter_Yield",      pd.Series(0)).sum()
        total_co2    = df_filtered.get("CO2_Avoided",         pd.Series(0)).sum()
        total_charge = df_filtered.get("Charge",              pd.Series(0)).sum()
        total_disch  = df_filtered.get("Discharge",           pd.Series(0)).sum()
        avg_peak     = df_filtered.get("Peak_Power",          pd.Series(0)).mean() or 0.0

        # ────────────────────────────────────────
        # GRID ENERGY INPUT
        # ────────────────────────────────────────
        st.subheader("Сүлжээнээс нийлүүлсэн эрчим хүч")
        grid_energy = st.number_input(
            "Сүлжээнээс нийлүүлсэн эрчим хүчийг оруулна уу (кВт·ц)",
            min_value=0.0,
            value=0.0,
            step=100.0,
            format="%.0f"
        )

        # ────────────────────────────────────────
        # KEY METRICS
        # ────────────────────────────────────────
        col1, col2, col3, col4, col5, col6 = st.columns(6)

        total_consumption = grid_energy + total_inv - total_charge + total_disch - 1200

        col1.metric("Үйлдвэрлэх боломжит эрчим хүч", f"{total_theo:,.0f} кВт·ц")
        col2.metric("Үйлдвэрлэсэн эрчим хүч",       f"{total_inv:,.0f} кВт·ц")
        col3.metric("CO₂ бууруулсан",                f"{total_co2:,.2f} тн")
        col4.metric("Нийт ЦЭХ хэрэглээ",            f"{total_consumption:,.0f} кВт·ц")
        col5.metric("Батарейнаас нийлүүлсэн",        f"{total_disch:,.0f} кВт·ц")
        col6.metric("Max чадлын дундаж",             f"{avg_peak:,.1f} кВт")

        st.markdown("---")

        # ────────────────────────────────────────
        # TABS
        # ────────────────────────────────────────
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Үйлдвэрлэл (харьцуулалт)",
            "Чадлын оргил утга",
            "CO₂ бууралт",
            "Батарей (цэнэглэлт / нийлүүлэлт)",
            "Сүлжээ vs Өөрийн үйлдвэрлэл"
        ])

        with tab1:
            st.subheader("Үйлдвэрлэх боломжит vs Үйлдвэрлэсэн эрчим хүч")
            fig1 = go.Figure()
            if "Theoretical_Yield" in df_filtered.columns:
                fig1.add_trace(go.Scatter(x=df_filtered["Date"], y=df_filtered["Theoretical_Yield"],
                                          name="Үйлдвэрлэх боломжит [кВт.ц]", line=dict(color='#10b981', width=2.5)))
            if "Inverter_Yield" in df_filtered.columns:
                fig1.add_trace(go.Scatter(x=df_filtered["Date"], y=df_filtered["Inverter_Yield"],
                                          name="Үйлдвэрлэсэн [кВт.ц]", line=dict(color='#3b82f6', width=2.5)))
            fig1.update_layout(height=500, hovermode="x unified",
                               legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig1, use_container_width=True)

        with tab2:
            st.subheader("Өдрийн хамгийн их чадал [кВт]")
            y_data = df_filtered.get("Peak_Power", pd.Series(0, index=df_filtered.index))
            fig2 = px.bar(df_filtered, x="Date", y=y_data, color_discrete_sequence=["#14b8a6"])
            fig2.update_layout(height=500)
            st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            st.subheader("CO₂ бууруулсан (тн)")
            y_data = df_filtered.get("CO2_Avoided", pd.Series(0, index=df_filtered.index))
            fig3 = px.line(df_filtered, x="Date", y=y_data, color_discrete_sequence=["#f59e0b"])
            fig3.update_traces(line=dict(width=2.8))
            fig3.update_layout(height=500)
            st.plotly_chart(fig3, use_container_width=True)

        with tab4:
            st.subheader("Батарей: Цэнэглэлт ба Нийлүүлэлт [кВт.ц]")
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(x=df_filtered["Date"], y=df_filtered.get("Charge", pd.Series(0)),
                                  name="Батарейг цэнэглэсэн", marker_color="#8b5cf6"))
            fig4.add_trace(go.Bar(x=df_filtered["Date"], y=df_filtered.get("Discharge", pd.Series(0)),
                                  name="Батарейнаас нийлүүлсэн", marker_color="#ec4899"))
            fig4.update_layout(barmode="group", height=500,
                               legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig4, use_container_width=True)

        with tab5:
            st.subheader("Сүлжээнээс нийлүүлсэн vs Өөрийн үйлдвэрлэсэн эрчим хүч")
            total_produced = df_filtered.get("Inverter_Yield", pd.Series(0)).sum()
            total_grid = grid_energy
            if total_produced + total_grid == 0:
                st.info("Өгөгдөл байхгүй эсвэл нийт 0 байна.")
            else:
                pie_data = pd.DataFrame({
                    "Төрөл": ["Өөрийн үйлдвэрлэсэн", "Сүлжээнээс нийлүүлсэн"],
                    "Эрчим хүч (кВт·ц)": [total_produced, total_grid]
                })
                fig_pie = px.pie(pie_data, values="Эрчим хүч (кВт·ц)", names="Төрөл",
                                 color_discrete_sequence=["#3b82f6", "#ef4444"], hole=0.4)
                fig_pie.update_traces(textposition='outside', textinfo='label+percent',
                                      insidetextorientation='horizontal', rotation=0, sort=False)
                fig_pie.update_layout(height=550,
                                      margin=dict(l=40, r=40, t=40, b=120),
                                      legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))
                st.plotly_chart(fig_pie, use_container_width=True)

        # ────────────────────────────────────────
        # DATA TABLE - safe formatting
        # ────────────────────────────────────────
        st.markdown("---")
        st.subheader("Гол үзүүлэлтүүд")

        display_df = df_filtered.copy()
        display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m-%d")

        possible_cols = ["Theoretical_Yield", "Inverter_Yield", "Peak_Power",
                         "CO2_Avoided", "Charge", "Discharge"]
        existing_cols = [c for c in possible_cols if c in display_df.columns]

        rename_dict_display = {
            "Date": "Огноо",
            "Theoretical_Yield": "Боломжит [кВт.ц]",
            "Inverter_Yield": "Үйлдвэрлэсэн [кВт.ц]",
            "Peak_Power": "Хамгийн их чадал [кВт]",
            "CO2_Avoided": "CO₂ бууруулсан [тн]",
            "Charge": "Цэнэглэсэн [кВт.ц]",
            "Discharge": "Нийлүүлсэн [кВт.ц]"
        }

        st.dataframe(
            display_df[["Date"] + existing_cols]
            .rename(columns=rename_dict_display),
            use_container_width=True,
            height=400,
            column_config={
                "Огноо": st.column_config.TextColumn(),
                "Боломжит [кВт.ц]": st.column_config.NumberColumn(format="%,.0f"),
                "Үйлдвэрлэсэн [кВт.ц]": st.column_config.NumberColumn(format="%,.0f"),
                "Хамгийн их чадал [кВт]": st.column_config.NumberColumn(format="%,.1f"),
                "CO₂ бууруулсан [тн]": st.column_config.NumberColumn(format="%,.2f"),
                "Цэнэглэсэн [кВт.ц]": st.column_config.NumberColumn(format="%,.0f"),
                "Нийлүүлсэн [кВт.ц]": st.column_config.NumberColumn(format="%,.0f"),
            }
        )

    except Exception as e:
        st.error(f"Файлыг уншихад алдаа гарлаа: {str(e)}")
        st.info("Шалгах зүйлс:\n• Файл .xlsx форматтай эсэх\n• Sheet1 байгаа эсэх\n• Эхний мөрүүд зөв эсэх (skiprows=1 тохиромжтой эсэх)")

else:
    st.info("Файл сонгоно уу")

st.markdown("---")
st.caption("МАК НЦС үйлдвэрлэл")
