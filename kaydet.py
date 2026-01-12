import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Excel'i oku (Header 3. satırda - İndeks 2)
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            target_col = 'Toplam Cam Zayi Oranı'
            
            # 2. Satır Seçimi (Excel 26-43 aralığı)
            final_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            # 3. HATA ÇÖZÜMÜ: Sayıya Zorla ve Temizle
            # 'Oran' veya 'Hedef' kelimesi geçen TÜM sütunları bul
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            
            for col in oran_cols:
                # errors='coerce' sayesinde yazı olan hücreler NaN (boş veri) olur ve HATA VERMEZ
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            # 4. SIRALAMA (Büyükten Küçüğe)
            if target_col in final_df.columns:
                final_df = final_df.sort_values(by=target_col, ascending=False)

            # 5. EN BAŞA BÖLGE
