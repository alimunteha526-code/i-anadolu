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

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            target_col = 'Toplam Cam Zayi Oranı'
            
            # 2. Satır Seçimi (Excel 26-43 aralığı)
            final_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            # 3. Veri Temizleme ve Sayıya Çevirme
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            for col in oran_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            # 4. Sıralama (Büyükten Küçüğe)
            if target_col in final_df.columns:
                final_df = final_df.sort_values(by=target_col, ascending=False)

            # 5. Başlık Satırı Oluşturma
            baslik_satiri = pd.DataFrame(columns=final_df.columns)
            baslik_satiri.loc[0] = [""] * len(final_df.columns)
            baslik_satiri.iloc[0, 0] = "İÇ ANADOLU"
            baslik_satiri.iloc[0, 1] = "BÖLGESİ"
            
            final_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            # 6. Güvenli Formatlama Fonksiyonu (Görünüm için)
            def format_yuzde(x):
                if pd.isna(x) or isinstance(x, str) or x == "": return x
                return "{:.1%}".format(x)

            # Görüntüleme için kopyasını oluştur (Excel'i bozmamak için)
            display_df = final_df.copy()
            for col in oran_cols:
                display_df[col] = display_df[col].apply(format_yuzde)

            # 7. Görsel Stil (Streamlit Önizleme)
            def stil_uygula(row):
                if row.iloc[0] == "İÇ ANADOLU":
                    return ['background-color: #2c3e50; color: white; font-weight: bold'] * len(row)
                return [''] * len(row)

            styled_df = display_df.style.apply(stil_uygula, axis=1)\
                .set_properties(**{'text-align': 'center', 'border': '1px solid black'})\
                .hide(axis="index")

            st.write("### Düzenlenmiş Liste Önizlemesi")
            st.write(styled_df)

            # --- İNDİRME BUTONLARI ---
            col1, col2 = st.columns(2)

            with col1:
                # EXCEL OLARAK İNDİR
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Zayi_Raporu')
                
                st.download_button(
                    label="📥 Excel Olarak İndir",
                    data=buffer_excel.getvalue(),
                    file_name="zayi_listesi_duzenlenmiş.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col2:
                # PDF OLARAK İNDİR (Basit Tablo Modu)
                pdf = FPDF(orientation='L', unit='mm', format='A4')
                pdf.add_page()
                pdf.set_font("Arial", size=8)
                
                # Tablo Genişliği Ayarı
                col_width = pdf.w / (len(final_df.columns) + 1)
                
                # Başlıkları Yaz
                pdf.set_fill_color(200, 200, 200)
                for col in final_df.columns:
                    pdf.cell(col_width, 10, str(col)[:10], border=1, fill=True)
                pdf.ln()

                # Verileri Yaz
                for i, row in display_df.iterrows():
                    if row.iloc[0] == "İÇ ANADOLU":
                        pdf.set_fill_color(44, 62, 80)
                        pdf.set_text_color(255, 255, 255)
                    else:
                        pdf.set_fill_color(255, 255, 255)
                        pdf.set_text_color(0, 0, 0)
                    
                    for val in row:
                        pdf.cell(col_width, 8, str(val)[:10], border=1, fill=True)
                    pdf.ln()

                pdf_output = pdf.output(dest='S').encode('latin-1', 'replace')
                st.download_button(
                    label="📥 PDF Olarak İndir",
                    data=pdf_output,
                    file_name="zayi_listesi.pdf",
                    mime="application/pdf"
                )

        else:
            st.error("'Üst Birim' sütunu bulunamadı!")
                
    except Exception as e:
        st.error(f"Hata: {e}")
