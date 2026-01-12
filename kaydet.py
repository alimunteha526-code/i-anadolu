import streamlit as st
import pandas as pd
import io
from fpdf import FPDF

# Sayfa Ayarları
st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            target_col = 'Toplam Cam Zayi Oranı'
            
            # Veri Hazırlama (26-43 aralığı)
            final_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            # Sayısal dönüşüm ve sıralama
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            for col in oran_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            if target_col in final_df.columns:
                final_df = final_df.sort_values(by=target_col, ascending=False)

            # Başlık Satırı
            baslik_satiri = pd.DataFrame(columns=final_df.columns)
            baslik_satiri.loc[0] = [""] * len(final_df.columns)
            baslik_satiri.iloc[0, 0] = "IC ANADOLU BOLGESI" # Türkçe karakter hatası riskine karşı
            
            report_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            st.write("### Düzenlenmiş Liste")
            st.dataframe(report_df)

            col1, col2 = st.columns(2)

            with col1:
                # EXCEL ÇIKTISI
                buffer_ex = io.BytesIO()
                with pd.ExcelWriter(buffer_ex, engine='xlsxwriter') as writer:
                    report_df.to_excel(writer, index=False)
                st.download_button("📥 Excel Olarak İndir", buffer_ex.getvalue(), "zayi_listesi.xlsx")

            with col2:
                # GELİŞMİŞ PDF ÇIKTISI
                try:
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 12)
                    pdf.cell(0, 10, "IC ANADOLU AEL ZAYI RAPORU", ln=True, align='C')
                    pdf.set_font("Helvetica", size=7) # 17 sütun için küçük font
                    
                    # Sütun Genişlikleri
                    col_width = pdf.epw / len(report_df.columns)
                    
                    # Tabloyu Çiz
                    for row in report_df.itertuples(index=False):
                        for item in row:
                            val = str(item) if pd.notna(item) else "-"
                            # Oranları % formatına çevir
                            if isinstance(item, float) and item < 1:
                                val = f"{item:.1%}"
                            pdf.cell(col_width, 8, val[:12], border=1)
                        pdf.ln()

                    pdf_bytes = pdf.output()
                    st.download_button("📥 PDF Olarak İndir", pdf_bytes, "zayi_raporu.pdf")
                except Exception as e:
                    st.error(f"PDF Oluşturulamadı: {e}")
        else:
            st.error("'Üst Birim' bulunamadı.")
    except Exception as e:
        st.error(f"Sistem Hatası: {e}")
