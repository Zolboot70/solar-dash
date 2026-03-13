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

# ────────────────────────────────────────
# TITLE & UPLOAD
# ────────────────────────────────────────
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

        rename_dict = {
            "Statistical Period": "Date",
            "Theoretical Yield (kWh)": "Theoretical_Yield",
            "Inverter Yield (kWh)": "Inverter_Yield",
            "Peak Power (kW)": "Peak_Power",
            "CO₂ Avoided (t)": "CO2_Avoided",
            "Charge (kWh)": "Charge",
            "Discharge (kWh)": "Discharge",
        }

        available_cols = [c for c in df.columns if any(k in c for k in rename_dict)]
        df = df[available_cols].copy()

        for old, new in rename_dict.items():
            for col in df.columns:
                if old in col:
                    df.rename(columns={col: new}, inplace=True)
                    break

        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
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
        total_theo   = df_filtered.get("Theoretical_Yield", pd.Series(0)).sum()
        total_inv    = df_filtered.get("Inverter_Yield", pd.Series(0)).sum()
        total_co2    = df_filtered.get("CO2_Avoided", pd.Series(0)).sum()
        total_charge = df_filtered.get("Charge", pd.Series(0)).sum()
        total_disch  = df_filtered.get("Discharge", pd.Series(0)).sum()
        avg_peak     = df_filtered.get("Peak_Power", pd.Series(0)).mean()
        avg_peak     = 0.0 if pd.isna(avg_peak) else avg_peak

        # ────────────────────────────────────────
        # GRID ENERGY INPUT
        # ────────────────────────────────────────
        st.subheader("Сүлжээнээс нийлүүлсэн эрчим хүч")
        grid_energy = st.number_input(
            "Сүлжээнээс нийлүүлсэн эрчим хүчийг оруулна уу (кВт·ц)",
            min_value=0.0,
            value=0.0,
            step=100.0,
            format="%.0f",
            help="Сүлжээнээс авсан нийт эрчим хүчийг оруулна уу"
        )

        # ────────────────────────────────────────
        # KEY METRICS
        # ────────────────────────────────────────
        col1, col2, col3, col4, col5, col6 = st.columns(6)

        total_consumption = (
            grid_energy +
            total_inv -
            total_charge +
            total_disch -
            1200
        )

        col1.metric("Үйлдвэрлэх боломжит эрчим хүч", f"{total_theo:,.0f} кВт·ц")
        col2.metric("Үйлдвэрлэсэн эрчим хүч",       f"{total_inv:,.0f} кВт·ц")
        col3.metric("CO₂ бууруулсан",                f"{total_co2:,.2f} тн")
        col4.metric("Нийт ЦЭХ хэрэглээ",            f"{total_consumption:,.0f} кВт·ц")
        col5.metric("Батарейнаас нийлүүлсэн",        f"{total_disch:,.0f} кВт·ц")
        col6.metric("Max чадлын дундаж",             f"{avg_peak:,.1f} кВт")

        st.markdown("---")

        # ────────────────────────────────────────
        # CHARTS TABS
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
            fig1.add_trace(go.Scatter(
                x=df_filtered["Date"],
                y=df_filtered.get("Theoretical_Yield", pd.Series(0)),
                name="Үйлдвэрлэх боломжит [кВт.ц]",
                line=dict(color='#10b981', width=2.5)
            ))
            fig1.add_trace(go.Scatter(
                x=df_filtered["Date"],
                y=df_filtered.get("Inverter_Yield", pd.Series(0)),
                name="Үйлдвэрлэсэн [кВт.ц]",
                line=dict(color='#3b82f6', width=2.5)
            ))
            fig1.update_layout(
                height=500,
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig1, use_container_width=True)

        with tab2:
            st.subheader("Өдрийн хамгийн их чадал [кВт]")
            fig2 = px.bar(
                df_filtered,
                x="Date",
                y="Peak_Power" if "Peak_Power" in df_filtered else pd.Series(0),
                color_discrete_sequence=["#14b8a6"]
            )
            fig2.update_layout(height=500)
            st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            st.subheader("CO₂ бууруулсан (тн)")
            fig3 = px.line(
                df_filtered,
                x="Date",
                y="CO2_Avoided" if "CO2_Avoided" in df_filtered else pd.Series(0),
                color_discrete_sequence=["#f59e0b"]
            )
            fig3.update_traces(line=dict(width=2.8))
            fig3.update_layout(height=500)
            st.plotly_chart(fig3, use_container_width=True)

        with tab4:
            st.subheader("Батарей: Цэнэглэлт ба Нийлүүлэлт [кВт.ц]")
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(
                x=df_filtered["Date"],
                y=df_filtered.get("Charge", pd.Series(0)),
                name="Батарейг цэнэглэсэн",
                marker_color="#8b5cf6"
            ))
            fig4.add_trace(go.Bar(
                x=df_filtered["Date"],
                y=df_filtered.get("Discharge", pd.Series(0)),
                name="Батарейнаас нийлүүлсэн",
                marker_color="#ec4899"
            ))
            fig4.update_layout(
                barmode="group",
                height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
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
                
                fig_pie = px.pie(
                    pie_data,
                    values="Эрчим хүч (кВт·ц)",
                    names="Төрөл",
                    color_discrete_sequence=["#3b82f6", "#ef4444"],
                    hole=0.4
                )
                
                fig_pie.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    insidetextorientation='horizontal',  # ← This prevents tilted/rotated text
                    rotation=0,
                    sort=False
                )
                
                fig_pie.update_layout(
                    height=600,
                    autosize=False,
                    margin=dict(l=40, r=40, t=60, b=140),
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.25,
                        xanchor="center",
                        x=0.5
                    ),
                    uniformtext_minsize=12,
                    uniformtext_mode='hide'
                )
                
                st.plotly_chart(fig_pie, use_container_width=True)

        # ────────────────────────────────────────
        # DATA TABLE
        # ────────────────────────────────────────
        st.markdown("---")
        st.subheader("Гол үзүүлэлтүүд")

        display_df = df_filtered.copy()
        display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m-%d")

        cols_to_show = ["Date"] + [c for c in ["Theoretical_Yield", "Inverter_Yield", "Peak_Power",
                                               "CO2_Avoided", "Charge", "Discharge"] if c in display_df]
        renamed = {
            "Date": "Огноо",
            "Theoretical_Yield": "Боломжит [кВт.ц]",
            "Inverter_Yield": "Үйлдвэрлэсэн [кВт.ц]",
            "Peak_Power": "Хамгийн их чадал [кВт]",
            "CO2_Avoided": "CO₂ бууруулсан [тн]",
            "Charge": "Цэнэглэсэн [кВт.ц]",
            "Discharge": "Нийлүүлсэн [кВт.ц]"
        }

        st.dataframe(
            display_df[cols_to_show].rename(columns=renamed).round(2),
            use_container_width=True,
            height=400
        )

    except Exception as e:
        st.error(f"Файлыг уншихад алдаа гарлаа:\n{str(e)}")
        st.info("Шалгах зүйлс:\n• .xlsx форматтай эсэх\n• Sheet1 байгаа эсэх\n• Загвар зөв эсэх")

else:
    st.info("Файл сонгоно уу")

st.markdown("---")
st.caption("МАК НЦС үйлдвэрлэл")
