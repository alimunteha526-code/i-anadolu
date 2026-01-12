import streamlit as st
import pandas as pd
import dataframe_image as dfi
import os

# Sayfa Ayarları
st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")

st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Excel'i oku (Veriler genellikle 3. satırdan başladığı için header=2)
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        # 2. Üst Birim sütununu bul ve 17 sütun sınırla
        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            
            # Excel 26-43 aralığını al (Python indeksi ile 22-40 arası)
            # 40. satırı (indeks 36) başlık için ayırıyoruz
            ust = df_full.iloc[22:36, start_col : start_col + 17].copy()
            alt = df_full.iloc[37:40, start_col : start_col + 17].copy()

            # 3. İç Anadolu Başlık Satırı Oluştur
            ara_satir = pd.DataFrame(columns=ust.columns)
            ara_satir.loc[0] = [""] * len(ust.columns)
            ara_satir.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ"

            final_df = pd.concat([ust, ara_satir, alt], ignore_index=True)

            # 4. Stil İşlemleri
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            zayi_col = 'Toplam Cam Zayi Oranı'

            def stil_ver(row):
                # Başlık satırını boya
                if row.iloc[0] == "İÇ ANADOLU BÖLGESİ":
                    return ['background-color: #2c3e50; color: white; font-weight: bold'] * len(row)
                
                # Normal satırlar için zayi oranına bak
                styles = [''] * len(row)
                if zayi_col in row.index:
                    val = row[zayi_col]
                    if isinstance(val, (int, float)) and val > 0.058:
                        idx = row.index.get_loc(zayi_col)
                        styles[idx] = 'background-color: #e74c3c; color: white'
                return styles

            # Tabloyu biçimlendir
            styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols if c in final_df.columns}, na_rep="")\
                .apply(stil_ver, axis=1)\
                .set_properties(**{'text-align': 'center', 'border': '1px solid black'})\
                .hide(axis="index")

            st.write("### Tablo Önizlemesi")
            st.dataframe(styled_df) # Önizlemede dataframe kullanmak daha güvenlidir

            # 5. FOTOĞRAF ÇIKTISI
            if st.button("🖼️ Fotoğrafı İndir"):
                with st.spinner('Lütfen bekleyin...'):
                    # Lokal bilgisayarda chrome hatası almamak için:
                    dfi.export(styled_df, "temp_rapor.png", table_conversion='chrome')
                    with open("temp_rapor.png", "rb") as f:
                        st.download_button("Resmi Kaydet", f, "zayi_listesi.png", "image/png")
        else:
            st.error("Excel'de 'Üst Birim' başlığı bulunamadı. Lütfen Excel'i kontrol edin.")
                
    except Exception as e:
        st.error(f"Hata detayı: {e}")
