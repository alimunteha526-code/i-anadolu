import streamlit as st
import pandas as pd
import io
from fpdf import FPDF

st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

# PDF için Türkçe karakter temizleyici
def tr_to_en(text):
    if not isinstance(text, str): return text
    mapping = {"İ": "I", "ı": "i", "Ş": "S", "ş": "s", "Ğ": "G", "ğ": "g", "Ç": "C", "ç": "c", "Ö": "O", "ö": "o", "Ü": "U", "ü": "u"}
    for tr, en in mapping.items(): text = text.replace(tr, en)
    return text

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        target_sort_col = 'Net Satış Miktarı (Cam)'
        zayi_oran_col = 'Toplam Cam Zayi Oranı'

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            # Veri aralığını al (26-43 satırları)
            report_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            # Sayısal dönüşüm
            numeric_cols = [c for c in report_df.columns if any(x in str(c) for x in ['Oran', 'Hedef', 'Miktarı', 'Adet'])]
            for col in numeric_cols:
                report_df[col] = pd.to_numeric(report_df[col], errors='coerce')

            # --- SIRALAMA (Net Satışa Göre En Yüksek En Üstte) ---
            if target_sort_col in report_df.columns:
                report_df = report_df.sort_values(by=target_sort_col, ascending=False)

            st.write(f"### {target_sort_col} Göre Sıralanmış Liste")
            st.dataframe(report_df)

            col1, col2 = st.columns(2)

            with col1:
                # --- PROFESYONEL EXCEL ÇIKTISI ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                    report_df.to_excel(writer, index=False, sheet_name='Zayi_Raporu')
                    workbook = writer.book
                    worksheet = writer.sheets['Zayi_Raporu']

                    # Formatlar
                    header_fmt = workbook.add_format({'bg_color': '#1f4e78', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'})
                    standard_fmt = workbook.add_format({'border': 1, 'align': 'center'})
                    percent_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})
                    critical_fmt = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center', 'num_format': '0.0%'})

                    # Başlıkları ve Sütunları Biçimlendir
                    for col_num, value in enumerate(report_df.columns):
                        worksheet.write(0, col_num, value, header_fmt)
                        worksheet.set_column(col_num, col_num, 18)

                    # Verileri Renklendir (Kritik Zayi Kontrolü)
                    for row_num in range(1, len(report_df) + 1):
                        row_data = report_df.iloc[row_num-1]
                        for col_num, val in enumerate(row_data):
                            col_name = report_df.columns[col_num]
                            
                            if col_name == zayi_oran_col and pd.notnull(val) and val > 0.058:
                                worksheet.write(row_num, col_num, val, critical_fmt)
                            elif 'Oran' in col_name or 'Hedef' in col_name:
                                worksheet.write(row_num, col_num, val, percent_fmt)
                            else:
                                worksheet.write(row_num, col_num, val if pd.notnull(val) else "", standard_fmt)

                st.download_button("📥 Renkli Excel İndir", output.getvalue(), "ic_anadolu_rapor.xlsx")

            with col2:
                # --- RENKLİ PDF ÇIKTISI ---
                try:
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 10)
                    pdf.cell(0, 10, "IC ANADOLU AEL SATIS VE ZAYI RAPORU", ln=True, align='C')
                    pdf.set_font("Helvetica", size=6)
                    col_width = pdf.epw / len(report_df.columns)

                    for i, row in report_df.iterrows():
                        for col_idx, item in enumerate(row):
                            col_name = report_df.columns[col_idx]
                            val = tr_to_en(str(item)) if pd.notna(item) else ""

                            # Kritik Zayi Renklendirmesi (Kırmızı)
                            if col_name == zayi_oran_col and isinstance(item, (float, int)) and item > 0.058:
                                pdf.set_fill_color(255, 0, 0)
                                pdf.set_text_color(255, 255, 255)
                            else:
                                pdf.set_fill_color(255, 255, 255)
                                pdf.set_text_color(0, 0, 0)

                            if isinstance(item, float) and 0 < item < 1: val = f"{item:.1%}"
                            pdf.cell(col_width, 7, val[:12], border=1, fill=True, align='C')
                        pdf.ln()

                    st.download_button("📥 Renkli PDF İndir", bytes(pdf.output()), "ic_anadolu_rapor.pdf")
                except Exception as e:
                    st.error(f"PDF Hatası: {e}")

    except Exception as e:
        st.error(f"Sistem Hatası: {e}")
