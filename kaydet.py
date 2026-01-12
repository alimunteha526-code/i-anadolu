import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Özel Aralık Raporu", layout="wide")

st.title("📊 Satır Aralığına Göre Mağaza Raporu")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    # 1. Excel'in tamamını oku (Ham haliyle)
    df = pd.read_excel(yuklenen_dosya, header=None)

    # 2. SATIR FİLTRELEME: 26 ile 43. satırlar (Python 0'dan başladığı için 25:43 yapılır)
    # Excel'deki 26. satır Python'da 25. indekstir.
    df_range = df.iloc[25:43].copy()

    # 3. Başlıkları Ayarla: Normalde başlıklar 3. satırda (indeks 2)
    basliklar = df.iloc[2].values
    df_range.columns = basliklar
    df_range.columns = df_range.columns.str.strip() # Boşlukları temizle

    if 'Üst Birim' in df_range.columns:
        # 4. SÜTUN FİLTRELEME: 'Üst Birim'den başla ve 17 sütun al
        ust_birim_idx = list(df_range.columns).index('Üst Birim')
        final_df = df_range.iloc[:, ust_birim_idx : ust_birim_idx + 17].copy()

        # 5. BİÇİMLENDİRME: Oranları % yap ve tabloyu daralt
        oran_sutunlari = [col for col in final_df.columns if 'Oran' in str(col) or 'Hedef' in str(col)]
        
        styled_df = final_df.style.format({col: "{:.1%}" for col in oran_sutunlari if col in final_df.columns})\
            .set_properties(**{
                'text-align': 'center',
                'font-size': '11px',
                'white-space': 'nowrap',
                'border': '1px solid lightgrey'
            })

        st.write("### 26-43. Satırlar Arası Rapor")
        st.write(styled_df)

        if st.button("🖼️ Fotoğraf Olarak İndir"):
            with st.spinner('Görüntü oluşturuluyor...'):
                resim_yolu = "ozel_aralik.png"
                dfi.export(styled_df, resim_yolu, table_conversion='chrome')
                
                with open(resim_yolu, "rb") as file:
                    st.download_button("Dosyayı Kaydet", file, "magaza_listesi.png", "image/png")
    else:
        st.error("'Üst Birim' sütunu bulunamadı. Lütfen satır aralığını kontrol edin.")
