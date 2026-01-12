import streamlit as st
import pandas as pd
import dataframe_image as dfi
import os

# Sayfa Ayarları
st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")

# ANA BAŞLIK
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Excel'i oku (Header 3. satırda kabul ediliyor - İndeks 2)
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        # 2. Üst Birim sütunundan itibaren 17 sütun al
        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            
            # Excel Satır 26'dan 43'e kadar olan aralığı seçiyoruz.
            # (Header=2 olduğu için df'in 0. satırı Excel'in 4. satırıdır. 
            # Bu yüzden Excel 26 -> df 22, Excel 43 -> df 40 olur)
            
            # Üst Bölüm (Excel 26-39 arası)
            ust_kisim = df_full.iloc[22:36, start_col : start_col + 17].copy()
            # Alt Bölüm (Excel 41-43 arası)
            alt_kisim = df_full.iloc[37:40, start_col : start_col + 17].copy()

            # 3. Ara Başlık Satırı (Excel 40. Satır yerine geçer)
            ara_satir = pd.DataFrame(columns=ust_kisim.columns)
            ara_satir.loc[0] = [""] * len(ust_kisim.columns)
            ara_satir.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ"

            # Hepsini birleştir
            final_df = pd.concat([ust_kisim, ara_satir, alt_kisim], ignore_index=True)

            # 4. GÖRSEL STİL FONKSİYONLARI
            def satir_stili(row):
                if row.iloc[0] == "İÇ ANADOLU BÖLGESİ":
                    return ['background-color: #2c3e50; color: white; font-weight: bold; text-align: center'] * len(row)
                return [''] * len(row)

            def oran_renklendir(v):
                if isinstance(v, (int, float)) and v > 0.058:
                    return 'background-color: #e74c3c; color: white; font-weight: bold'
                return ''

            # 5. TABLO FORMATLAMA
            oran_cols = [c for c in final_df.columns if 'Oran' in c or 'Hedef' in c]
            target_col = 'Toplam Cam Zayi Oranı'

            styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols}, na_rep="")\
                .apply(satir_stili, axis=1)\
                .applymap(oran_renklendir, subset=[target_col] if target_col in final_df.columns else [])\
                .set_properties(**{
                    'text-align': 'center',
                    'font-size': '12px',
                    'border': '1px solid black',
                    'white-space': 'nowrap'
                })\
                .hide(axis="index")

            st.write("### Düzenlenmiş Liste Önizlemesi (26-43. Satırlar)")
            st.write(styled_df)

            # 6. FOTOĞRAF ÇIKTISI
            if st.button("🖼️ Fotoğrafı Hazırla ve İndir"):
                with st.spinner('Görsel oluşturuluyor...'):
                    resim = "zayi_raporu.png"
                    dfi.export(styled_df, resim, table_conversion='chrome')
                    
                    with open(resim, "rb") as f:
                        st.download_button("Görseli Kaydet", f, "zayi_listesi.png", "image/png")
        else:
            st.error("Excel'de 'Üst Birim' sütunu bulunamadı!")
                
    except Exception as e:
        st.error(f"Bir hata oluştu: {e}")
