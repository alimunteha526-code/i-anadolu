import streamlit as st
import pandas as pd
import io
from fpdf import FPDF # fpdf2 kütüphanesi önerilir

st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Veriyi Oku ve Hazırla
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            target_col = 'Toplam Cam Zayi Oranı'
            
            # Satır Seçimi (Excel 26-43 -> İndeks 22-40)
            final_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            # Veri Temizleme (Sayıya Çevirme)
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            for col in oran_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            # Sıralama
            if target_col in final_df.columns:
                final_df = final_df.sort_values(by=target_col, ascending=False)

            # Başlık Satırı Ekleme
            baslik_satiri = pd.DataFrame(columns=final_df.columns)
            baslik_satiri.loc[0] = [""] * len(final_df.columns)
            baslik_satiri.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ"
            final_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            # --- ÖNİZLEME ---
            st.write("### Liste Önizlemesi")
            st.dataframe(final_df)

            # --- İNDİRME ALANI ---
            col1, col2 = st.columns(2)

            with col1:
                # EXCEL ÇIKTISI
                buffer_ex = io.BytesIO()
                with pd.ExcelWriter(buffer_ex, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Zayi_Raporu')
                st.download_button("📥 Excel Olarak İndir", buffer_ex.getvalue(), "zayi_listesi.xlsx")

            with col2:
                # PDF ÇIKTISI (Güvenli Mod)
                try:
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    pdf.set_font("Helvetica", size=10) # Standart font hata riskini azaltır
                    
                    # Tabloyu PDF'e dök
                    pdf.cell(0, 10, "IC ANADOLU AEL ZAYI LISTESI", ln=True, align='C')
                    pdf.ln(5)
                    
                    # Verileri satır satır ekle (Sadece ilk 5 sütunu örnek alıyoruz sığması için)
                    for i, row in final_df.head(20).iterrows():
                        text_row = " | ".join([str(val)[:15] for val in row[:5]])
                        pdf.cell(0, 8, text_row, border=1, ln=True)
                    
                    pdf_output = pdf.output(dest='S')
                    st.download_button("📥 PDF Olarak İndir (Özet)", pdf_output, "zayi_raporu.pdf")
                except Exception as pdf_error:
                    st.warning("PDF oluşturulurken bir kütüphane sorunu oluştu, lütfen Excel indirin.")

        else:
            st.error("'Üst Birim' sütunu bulunamadı!")
                
    except Exception as e:
        st.error(f"Hata: {e}")
