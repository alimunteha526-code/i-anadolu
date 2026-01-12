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
        # 1. Excel'i oku
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            target_col = 'Toplam Cam Zayi Oranı'
            
            # Veri Hazırlama (Excel 26-43 aralığı)
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
            baslik_satiri.iloc[0, 0] = "IC ANADOLU BOLGESI"
            
            report_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            st.write("### Düzenlenmiş Liste")
            st.dataframe(report_df)

            col1, col2 = st.columns(2)

            with col1:
                # --- EXCEL ÇIKTISI ---
                buffer_ex = io.BytesIO()
                with pd.ExcelWriter(buffer_ex, engine='xlsxwriter') as writer:
                    report_df.to_excel(writer, index=False)
                st.download_button(
                    label="📥 Excel Olarak İndir",
                    data=buffer_ex.getvalue(),
                    file_name="ic_anadolu_zayi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col2:
                # --- PDF ÇIKTISI (HATA GİDERİLMİŞ) ---
                try:
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 11)
                    pdf.cell(0, 10, "IC ANADOLU AEL ZAYI RAPORU", ln=True, align='C')
                    pdf.set_font("Helvetica", size=6) # 17 sütun için fontu küçülttük
                    
                    # Sayfa genişliğini sütun sayısına böl
                    col_width = pdf.epw / len(report_df.columns)
                    
                    for i, row in report_df.iterrows():
                        # Bölge başlığı satırı için renk değiştir (isteğe bağlı)
                        if row.iloc[0] == "IC ANADOLU BOLGESI":
                            pdf.set_fill_color(44, 62, 80)
                            pdf.set_text_color(255, 255, 255)
                        else:
                            pdf.set_fill_color(255, 255, 255)
                            pdf.set_text_color(0, 0, 0)
                            
                        for item in row:
                            val = str(item) if pd.notna(item) else ""
                            # Sayısal oranları görselleştir
                            if isinstance(item, (float, int)) and 0 < item < 1:
                                val = f"{item:.1%}"
                            
                            pdf.cell(col_width, 7, val[:12], border=1, fill=True)
                        pdf.ln()

                    # KRİTİK DÜZELTME: bytearray'i bytes formatına çeviriyoruz
                    pdf_output = pdf.output()
                    pdf_bytes = bytes(pdf_output) 
                    
                    st.download_button(
                        label="📥 PDF Olarak İndir",
                        data=pdf_bytes,
                        file_name="ic_anadolu_zayi.pdf",
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"PDF Hatası: {e}")
        else:
            st.error("'Üst Birim' bulunamadı.")
    except Exception as e:
        st.error(f"Sistem Hatası: {e}")
