import streamlit as st
import pandas as pd
import dataframe_image as dfi

# Sayfa ayarları
st.set_page_config(page_title="Mağaza Raporu", layout="wide")

# Senin belirttiğin mağaza kodları
varsayilan_magazalar = ["M38001", "M38003", "M38002", "M38005", "M38004", "M42001"]

st.title("📊 Mağaza Cam Zayi Raporu")

# Dosya yükleme
yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Excel'i oku (İlk 2 satırı atla)
        df = pd.read_excel(yuklenen_dosya, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns] # Sütun isimlerini temizle

        if 'Üst Birim' in df.columns:
            # 2. Mağaza kodlarına göre filtrele
            final_df = df[df['Üst Birim'].isin(varsayilan_magazalar)].copy()
            
            # 3. 'Üst Birim'den itibaren 17 sütun al
            ust_birim_idx = final_df.columns.get_loc('Üst Birim')
            final_df = final_df.iloc[:, ust_birim_idx : ust_birim_idx + 17]

            # 4. Oranları % yap ve %5.8 üzerini kırmızı boya
            oran_cols = [c for c in final_df.columns if 'Oran' in c or 'Hedef' in c]
            
            def kirmizi_boya(val):
                if isinstance(val, (int, float)) and val > 0.058:
                    return 'background-color: #ffcccc; color: #cc0000; font-weight: bold'
                return ''

            styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols})\
                .applymap(kirmizi_boya, subset=[c for c in final_df.columns if 'Toplam Cam Zayi Oranı' in c])\
                .set_properties(**{
                    'text-align': 'center',
                    'font-size': '12px',
                    'border': '1px solid #ddd'
                })\
                .hide(axis="index")

            st.write("### Tablo Önizlemesi")
            st.write(styled_df)

            # 5. Fotoğraf İndirme
            if st.button("🖼️ Fotoğraf Olarak İndir"):
                resim_yolu = "rapor_cikti.png"
                dfi.export(styled_df, resim_yolu)
                with open(resim_yolu, "rb") as f:
                    st.download_button("Resmi Kaydet", f, "rapor.png", "image/png")
        else:
            st.error("'Üst Birim' sütunu bulunamadı!")
            
    except Exception as e:
        st.error(f"Hata oluştu: {e}")
