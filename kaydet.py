import streamlit as st
import pandas as pd
import io
from fpdf import FPDF

st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Excel'i oku
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        target_sort_col = 'Net Satış Miktarı (Cam)'
        zayi_oran_col = 'Toplam Cam Zayi Oranı'

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            
            # 2. Veri Aralığını Al (Excel 26-43)
            final_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            # 3. Veri Temizleme ve Sıralama
            numeric_cols = [c for c in final_df.columns if any(x in str(c) for x in ['Oran', 'Hedef', 'Miktarı', 'Adet'])]
            for col in numeric_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            # Net Satışa Göre Büyükten Küçüğe Sırala
            if target_sort_col in final_df.columns:
                final_df = final_df.sort_values(by=target_sort_col, ascending=False)

            # 4. Başlık Satırı Ekleme (Renklendirme için işaretliyoruz)
            baslik_satiri = pd.DataFrame(columns=final_df.columns)
            baslik_satiri.loc[0] = [""] * len(final_df.columns)
            baslik_satiri.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ"
            
            report_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            # --- EKRAN ÖNİZLEMESİ (Biçimlendirilmiş) ---
            def highlight_rows(row):
                if row.iloc[0] == "İÇ ANADOLU BÖLGESİ":
                    return ['background-color: #1f4e78; color: white; font-weight: bold'] * len(row)
                
                styles = [''] * len(row)
                if zayi_oran_col in row.index:
                    val = row[zayi_oran_col]
                    # %5.8 kritik eşik (0.058)
                    if pd.notnull(val) and isinstance(val, (float, int)) and val > 0.058:
                        idx = row.index.get_loc(zayi_oran_col)
                        styles[idx] = 'background-color: #ff0000; color: white; font-weight: bold'
                return styles

            st.write("### Excel Formatlı Önizleme")
            st.dataframe(report_df.style.apply(highlight_rows, axis=1).format(lambda x: f"{x:.1%}" if isinstance(x, float) and 0 < x < 1 else x))

            col1, col2 = st.columns(2)

            with col1:
                # --- PROFESYONEL EXCEL ÇIKTISI ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    report_df.to_excel(writer, index=False, sheet_name='Zayi_Raporu')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Zayi_Raporu']
                    
                    # Formatlar
                    header_fmt = workbook.add_format({'bg_color': '#1f4e78', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'})
                    percent_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})
                    standard_fmt = workbook.add_format({'border': 1, 'align': 'center'})
                    critical_fmt = workbook.add_format({'bg_color': '#ff0000', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center', 'num_format': '0.0%'})

                    # Sütunları ve Başlıkları Biçimlendir
                    for col_num, value in enumerate(report_df.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                        worksheet.set_column(col_num, col_num, 15)

                    # Verileri Biçimlendir
                    for row_num in range(1, len(report_df) + 1):
                        row_data = report_df.iloc[row_num-1]
                        for col_num, val in enumerate(row_data):
                            col_name = report_df.columns[col_num]
                            
                            # Özel Bölge Başlığı Formatı
                            if row_data.iloc[0] == "İÇ ANADOLU BÖLGESİ":
                                worksheet.write(row_num, col_num, val, header_fmt)
                            # Zayi Oranı Kritik Formatı
                            elif col_name == zayi_oran_col and pd.notnull(val) and val > 0.058:
                                worksheet.write(row_num, col_num, val, critical_fmt)
                            # Genel Yüzde Formatı
                            elif 'Oran' in col_name or 'Hedef' in col_name:
                                worksheet.write(row_num, col_num, val, percent_fmt)
                            else:
                                worksheet.write(row_num, col_num, val, standard_fmt)

                st.download_button("📥 Renkli Excel İndir", output.getvalue(), "ic_anadolu_renkli.xlsx")

            with col2:
                # --- PROFESYONEL PDF ÇIKTISI ---
                try:
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 10)
                    pdf.cell(0, 10, "IC ANADOLU BOLGESI RAPORU", ln=True, align='C')
                    pdf.set_font("Helvetica", size=6)
                    
                    col_width = pdf.epw / len(report_df.columns)
                    
                    for i, row in report_df.iterrows():
                        is_header = row.iloc[0] == "İÇ ANADOLU BÖLGESİ"
                        
                        for col_idx, item in enumerate(row):
                            col_name = report_df.columns[col_idx]
                            val = str(item) if pd.notna(item) else ""
                            
                            # PDF Renklendirme Mantığı
                            if is_header:
                                pdf.set_fill_color(31, 78, 120) # Lacivert
                                pdf.set_text_color(255, 255, 255)
                            elif col_name == zayi_oran_col and isinstance(item, (float, int)) and item > 0.058:
                                pdf.set_fill_color(255, 0, 0) # Kırmızı
                                pdf.set_text_color(255, 255, 255)
                            else:
                                pdf.set_fill_color(255, 255, 255)
                                pdf.set_text_color(0, 0, 0)
                            
                            if isinstance(item, float) and 0 < item < 1:
                                val = f"{item:.1%}"
                            
                            pdf.cell(col_width, 7, val[:12], border=1, fill=True, align='C')
                        pdf.ln()

                    pdf_bytes = bytes(pdf.output())
                    st.download_button("📥 Renkli PDF İndir", pdf_bytes, "ic_anadolu_renkli.pdf")
                except Exception as e:
                    st.error(f"PDF Hatası: {e}")

    except Exception as e:
        st.error(f"Sistem Hatası: {e}")
