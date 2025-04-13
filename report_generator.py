import pandas as pd
import matplotlib.pyplot as plt
import io

def process_excel(file_path):
    df = pd.read_excel(file_path)
    df['Date'] = pd.to_datetime(df['Date'])
    df['Day'] = df['Date'].dt.date

    total_days = df['Day'].nunique()
    total_employees = df['Employee'].nunique()
    total_entries = len(df)

    summary = f"📊 Umumiy statistika:\n- Unikal kunlar: {total_days}\n- Xodimlar soni: {total_employees}\n- Jami kirishlar: {total_entries}"

    # Hisobot uchun Excel fayl yaratish
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Ma\'lumotlar')
        summary_df = pd.DataFrame({
            'Ko\'rsatkich': ['Unikal kunlar', 'Xodimlar soni', 'Kirishlar soni'],
            'Qiymat': [total_days, total_employees, total_entries]
        })
        summary_df.to_excel(writer, index=False, sheet_name='Statistika')

    output.seek(0)

    # Diagramma rasmga olish
    chart_image = io.BytesIO()
    plt.figure(figsize=(6, 4))
    employee_counts = df['Employee'].value_counts()
    employee_counts.plot(kind='bar', color='skyblue')
    plt.title('Xodimlar bo\'yicha tashriflar soni')
    plt.xlabel('Xodim')
    plt.ylabel('Tashriflar soni')
    plt.tight_layout()
    plt.savefig(chart_image, format='png')
    chart_image.seek(0)

    return summary, output, chart_image
