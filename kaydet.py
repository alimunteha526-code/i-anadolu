import streamlit as st
import pandas as pd
import dataframe_image as dfi
import os

st.set_page_config(page_title="Rapor Oluşturucu", layout="wide")

st.title("📊 Cam Zayi Raporu (26-43. Satırlar)")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Excel'i oku (Başlıkları 3. satırdan al - Python indeksi 2)
        df = pd.read_excel(yuklenen_dosya, header=2)
        
        # 2. Excel'deki 26-43 arası satırları al 
        # (header=2 dediğimiz için df 3. satırdan başlar, 26. satır df içinde 23. indekse düşer)
        # Eğer bu aralık kayarsa alttaki rakamları (23, 41) 1-2 sayı artırıp azaltabilirsin.
        final_df = df.iloc[23:41, 2:19].copy() # 2. sütundan (Üst Birim) başla, 17 sütun git
        
        # Sütun isimlerini temizle
        final_df.columns = [str(c).strip() for c in final_df.columns]

        # 3. Yüzde Biçimlendirme ve Renklendirme Fonksiyonu
        def stil_uygula(v):
            if isinstance(v, (int, float)) and v > 0.058:
                return 'background-color: #e74c3c; color: white; font-weight: bold'
            return ''

        # Oran sütunlarını bul (% işareti eklemek için)
        oran_cols = [c for c in final_df.columns if 'Oran' in c or 'Hedef' in c]

        styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols})\
            .applymap(stil_uygula, subset=[c for c in final_df.columns if 'Toplam Cam Zayi Oranı' in c])\
            .set_properties(**{
                'text-align': 'center',
                'font-size': '12px',
                'border': '1px solid black',
                'white-space': 'nowrap'
            })\
            .set_caption("İÇ ANADOLU BÖLGESİ")\
            .hide(axis="index")

        st.write("### Tablo Önizlemesi")
        st.write(styled_df)

        if st.button("🖼️ Fotoğraf Olarak İndir"):
            resim = "rapor.png"
            # Fotoğraf oluşturma motoru (chrome yüklü olmalı)
            dfi.export(styled_df, resim, table_conversion='chrome')
            
            with open(resim, "rb") as f:
                st.download_button("Resmi Kaydet", f, "rapor.png", "image/png")
                
    except Exception as e:
        st.error(f"Bir hata oluştu: {e}")
        st.info("Lütfen yüklediğiniz Excel'in formatını kontrol edin.")
