import streamlit as st
import pandas as pd
import dataframe_image as dfi

# Sayfa ayarlarını geniş yapalım
st.set_page_config(page_title="Mağaza Raporu", layout="wide")

varsayilan_magazalar = ["M38001", "M38003", "M38002", "M38005", "M38004", "M42001"]

st.title("📊 Kompakt Mağaza Raporu")

secilen_kodlar = st.multiselect("Mağazaları Seçin:", options=varsayilan_magazalar, default=varsayilan_magazalar)

yuklenen_dosya = st.file_uploader("Excel Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    df = pd.read_excel(yuklenen_dosya, skiprows=2)
    df.columns = df.columns.str.strip()

    if 'Üst Birim' in df.columns:
        ust_birim_index = df.columns.get_loc('Üst Birim')
        final_df = df.iloc[:, ust_birim_index : ust_birim_index + 17].copy()
        final_df = final_df[final_df['Üst Birim'].isin(secilen_kodlar)]

        if not final_df.empty:
            # --- SIKIŞTIRMA VE BİÇİMLENDİRME ---
            oran_sutunlari = [col for col in final_df.columns if 'Oran' in col or 'Hedef' in col]
            
            # Stil ayarları: Yazı boyutu, hizalama ve boşlukları sıfırlama
            styled_df = final_df.style.format({col: "{:.1%}" for col in oran_sutunlari})\
                .set_properties(**{
                    'text-align': 'center', # Yazıları ortala
                    'font-size': '12px',    # Yazı boyutunu hafif küçült
                    'white-space': 'nowrap' # Yazıların alt satıra geçmesini engelle (sütunu daraltır)
                })\
                .set_table_styles([
                    {'selector': 'th', 'props': [('font-size', '12px'), ('text-align', 'center')]}
                ])
            
            st.write("### Önizleme (Daraltılmış)")
            st.write(styled_df)

            if st.button("🖼️ Fotoğrafı Al"):
                with st.spinner('Fotoğraf hazırlanıyor...'):
                    resim_yolu = "dar_tablo.png"
                    # 'chrome' modu sütunları en dar haline getirir
                    dfi.export(styled_df, resim_yolu, table_conversion='chrome')
                    
                    with open(resim_yolu, "rb") as file:
                        st.download_button("İndir", file, "magaza_rapor.png", "image/png")
