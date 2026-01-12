import streamlit as st
import pandas as pd
import dataframe_image as dfi

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
            
            # Satır Seçimi (Excel 26-43 aralığı)
            ust = df_full.iloc[22:36, start_col : start_col + 17].copy()
            alt = df_full.iloc[37:40, start_col : start_col + 17].copy()

            # 2. Ara Başlık Satırı
            ara_satir = pd.DataFrame(columns=ust.columns)
            ara_satir.loc[0] = [""] * len(ust.columns)
            ara_satir.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ"

            final_df = pd.concat([ust, ara_satir, alt], ignore_index=True)

            # 3. VERİ TEMİZLEME (HATA ÇÖZÜMÜ)
            # Oran içeren sütunlardaki metinleri sayıya çevir, çevrilemeyenleri NaN yap
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            for col in oran_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            # 4. STİL FONKSİYONLARI
            def stil_ver(row):
                if row.iloc[0] == "İÇ ANADOLU BÖLGESİ":
                    return ['background-color: #2c3e50; color: white; font-weight: bold'] * len(row)
                
                styles = [''] * len(row)
                zayi_col = 'Toplam Cam Zayi Oranı'
                if zayi_col in row.index:
                    val = row[zayi_col]
                    # Sadece sayı olanları ve %5.8'den büyük olanları boya
                    if pd.notnull(val) and isinstance(val, (int, float)) and val > 0.058:
                        idx = row.index.get_loc(zayi_col)
                        styles[idx] = 'background-color: #e74c3c; color: white'
                return styles

            # 5. TABLO GÖRÜNÜMÜ
            # Sayı olanlara yüzde formatı uygula, olmayanları boş bırak
            styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols}, na_rep="-")\
                .apply(stil_ver, axis=1)\
                .set_properties(**{'text-align': 'center', 'border': '1px solid black', 'white-space': 'nowrap'})\
                .hide(axis="index")

            st.write("### Tablo Önizlemesi")
            st.write(styled_df)

            if st.button("🖼️ Fotoğrafı İndir"):
                with st.spinner('Görsel hazırlanıyor...'):
                    resim = "zayi_listesi.png"
                    dfi.export(styled_df, resim, table_conversion='chrome')
                    with open(resim, "rb") as f:
                        st.download_button("Resmi Kaydet", f, "zayi_listesi.png", "image/png")
        else:
            st.error("'Üst Birim' sütunu bulunamadı.")
                
    except Exception as e:
        st.error(f"Hata detayı: {e}")
