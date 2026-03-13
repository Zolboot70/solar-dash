# solar_dashboard.py

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(
    page_title="Нарны цахилгаан станцын үйлдвэрлэл",
    layout="wide"
)

# ────────────────────────────────────────
#               TITLE & UPLOAD
# ────────────────────────────────────────
st.title("Нарны цахилгаан станцын үйлдвэрлэл")
st.markdown("PVMS-ын стандарт хэлбэртэй файл оруулна уу")

uploaded_file = st.file_uploader(
    "НЦС-ын үйлдвэрлэлийн мэдээг оруулах",
    type=["xlsx"],
    help="Файл дотор 'Sheet1' байх ёстой"
)

if uploaded_file is not None:
    try:
        # ────────────────────────────────────────
        #               READ DATA
        # ────────────────────────────────────────
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", skiprows=1)

        # Rename columns to English for easier coding (but we display Mongolian)
        rename_dict = {
            "Statistical Period": "Date",
            "Theoretical Yield (kWh)": "Theoretical_Yield",
            "Inverter Yield (kWh)": "Inverter_Yield",
            "Peak Power (kW)": "Peak_Power",
            "CO₂ Avoided (t)": "CO2_Avoided",
            "Charge (kWh)": "Charge",
            "Discharge (kWh)": "Discharge",
        }

        # Keep only columns we care about + try to match
        available_cols = [c for c in df.columns if any(k in c for k in rename_dict)]
        df = df[available_cols].copy()

        # Rename to standardized names
        for old, new in rename_dict.items():
            for col in df.columns:
                if old in col:
                    df.rename(columns={col: new}, inplace=True)
                    break

        # Convert Date
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])           # drop bad dates
        df = df.sort_values("Date")

        # ────────────────────────────────────────
        #               SIDEBAR FILTER
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

        # Filter dataframe
        if len(date_range) == 2:
            start_date, end_date = date_range
            mask = (df["Date"].dt.date >= start_date) & (df["Date"].dt.date <= end_date)
            df_filtered = df.loc[mask].copy()
        else:
            df_filtered = df.copy()

        # ────────────────────────────────────────
        #               KEY METRICS
        # ────────────────────────────────────────
        col1, col2, col3, col4, col5, col6 = st.columns(6)

        total_theo   = df_filtered["Theoretical_Yield"].sum()
        total_inv    = df_filtered["Inverter_Yield"].sum()
        total_co2    = df_filtered["CO2_Avoided"].sum()
        total_charge = df_filtered["Charge"].sum()
        total_disch  = df_filtered["Discharge"].sum()
        avg_peak     = df_filtered["Peak_Power"].mean()

        col1.metric("Үйлдвэрлэх боломжит эрчим хүч", f"{total_theo:,.0f} кВт·ц")
        col2.metric("Үйлдвэрлэсэн эрчим хүч",       f"{total_inv:,.0f} кВт·ц")
        col3.metric("CO₂ бууруулсан",                f"{total_co2:,.2f} тн")
        col4.metric("Батарейг цэнэглэсэн",           f"{total_charge:,.0f} кВт·ц")
        col5.metric("Батарейнаас нийлүүлсэн",        f"{total_disch:,.0f} кВт·ц")
        col6.metric("Max чадлын дундаж",     f"{avg_peak:,.1f} кВт")

        st.markdown("---")

        # ────────────────────────────────────────
        #               CHARTS
        # ────────────────────────────────────────
        tab1, tab2, tab3, tab4 = st.tabs([
            "Үйлдвэрлэл (харьцуулалт)",
            "Чадлын оргил утга",
            "CO₂ бууралт",
            "Батарей (цэнэглэлт / нийлүүлэлт)"
        ])

        with tab1:
            st.subheader("Үйлдвэрлэх боломжит vs Үйлдвэрлэсэн эрчим хүч")
            fig1 = go.Figure()

            fig1.add_trace(go.Scatter(
                x=df_filtered["Date"],
                y=df_filtered["Theoretical_Yield"],
                name="Үйлдвэрлэх боломжит эрчим хүч [кВт.ц]",
                line=dict(color='#10b981', width=2.5)
            ))

            fig1.add_trace(go.Scatter(
                x=df_filtered["Date"],
                y=df_filtered["Inverter_Yield"],
                name="Үйлдвэрлэсэн эрчим хүч [кВт.ц]",
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
                y="Peak_Power",
                title="",
                color_discrete_sequence=["#14b8a6"]
            )
            fig2.update_layout(height=500)
            st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            st.subheader("CO₂ бууруулсан (тн)")
            fig3 = px.line(
                df_filtered,
                x="Date",
                y="CO2_Avoided",
                title="",
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
                y=df_filtered["Charge"],
                name="Батарейг цэнэглэсэн",
                marker_color="#8b5cf6"
            ))

            fig4.add_trace(go.Bar(
                x=df_filtered["Date"],
                y=df_filtered["Discharge"],
                name="Батарейнаас нийлүүлсэн",
                marker_color="#ec4899"
            ))

            fig4.update_layout(
                barmode="group",
                height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig4, use_container_width=True)

        # ────────────────────────────────────────
        #               DATA TABLE
        # ────────────────────────────────────────
        st.markdown("---")
        st.subheader("Гол үзүүлэлтүүд")

        display_df = df_filtered.copy()
        display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m-%d")

        st.dataframe(
            display_df[["Date", "Theoretical_Yield", "Inverter_Yield",
                        "Peak_Power", "CO2_Avoided", "Charge", "Discharge"]]
            .rename(columns={
                "Date": "Огноо",
                "Theoretical_Yield": "Боломжит [кВт.ц]",
                "Inverter_Yield": "Үйлдвэрлэсэн [кВт.ц]",
                "Peak_Power": "Хамгийн их чадал [кВт]",
                "CO2_Avoided": "CO₂ бууруулсан [тн]",
                "Charge": "Цэнэглэсэн [кВт.ц]",
                "Discharge": "Нийлүүлсэн [кВт.ц]"
            })
            .round(2),
            use_container_width=True,
            height=400
        )

    except Exception as e:
        st.error(f"Файлыг уншихад алдаа гарлаа:\n{str(e)}")
        st.info("Дараах зүйлсийг шалгаарай:\n• Файл .xlsx форматтай эсэх\n• Sheet1 нэртэй хуудас байгаа эсэх\n• Загвар өмнөх жишээтэй төстэй эсэх")

else:
    st.info("Файл сонгоно уу ")

st.markdown("---")
st.caption("МАК НЦС үйлдювэрлэл")