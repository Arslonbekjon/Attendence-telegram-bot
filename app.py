import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Attendance Analyzer", layout="wide")
st.title("📊 Attendance Analyzer")

# Fayl yuklash
uploaded_file = st.file_uploader("📁 Excel faylni yuklang", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Sanani formatlash
    df['DateTime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str))
    df = df.sort_values(by=['Employee', 'DateTime'])

    sessions = []

    for employee, group in df.groupby('Employee'):
        group = group.reset_index(drop=True)
        i = 0
        while i < len(group) - 1:
            start = group.loc[i, 'DateTime']
            end = group.loc[i + 1, 'DateTime']
            date = group.loc[i, 'Date']
            duration = end - start
            sessions.append({
                'Employee': employee,
                'Date': date,
                'Start': start.time(),
                'End': end.time(),
                'Duration (minutes)': round(duration.total_seconds() / 60, 2)
            })
            i += 2

    session_df = pd.DataFrame(sessions)

    # Xodim bo‘yicha filter
    employees = session_df['Employee'].unique()
    selected_employee = st.selectbox("👤 Xodimni tanlang", ["Barchasi"] + list(employees))
    if selected_employee != "Barchasi":
        session_df = session_df[session_df['Employee'] == selected_employee]

    st.subheader("📋 Sessiyalar Jadvali")
    st.dataframe(session_df, use_container_width=True)

    # Umumiy davomiylik bo‘yicha eng yaxshi/yomon xodimlar
    total_duration = session_df.groupby('Employee')['Duration (minutes)'].sum()

    if not total_duration.empty:
        best_employee = total_duration.idxmax()
        best_duration = total_duration.max()

        worst_employee = total_duration.idxmin()
        worst_duration = total_duration.min()

        with st.expander("🏆 Eng yaxshi va eng yomon xodimlar", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.success(f"🔝 Eng ko‘p ishlagan: **{best_employee}** — {round(best_duration, 2)} daqiqa")
            with col2:
                st.error(f"⬇️ Eng kam ishlagan: **{worst_employee}** — {round(worst_duration, 2)} daqiqa")

    # Bar chart – umumiy ish vaqti
    st.subheader("⏱️ Umumiy ish vaqti (daqiqa hisobida)")
    duration_by_employee = session_df.groupby('Employee')['Duration (minutes)'].sum()
    st.bar_chart(duration_by_employee)

    # Excel yuklab olish
    @st.cache_data
    def convert_df(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    excel_data = convert_df(session_df)

    st.download_button(
        label="⬇️ Excel faylni yuklab olish",
        data=excel_data,
        file_name="session_data.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

else:
    st.info("Iltimos, Excel fayl yuklang.")
