import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Özel Satır Raporu", layout="wide")

st.title("📊 Satır 26-43 Analiz Raporu")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    # 1. Excel'i ham veri olarak oku (hiçbir satırı atlama)
    df_raw = pd.read_excel(yuklenen_dosya, header=None)

    # 2. SATIR SEÇİMİ: Excel'deki 26-43 aralığı (İndeks olarak 25'ten 43'e kadar)
    # 25:43 yazınca 25 dahil, 43 dahil değildir (yani tam 18 satır alır)
    df_range = df_raw.iloc[25:43].copy()

    # 3. BAŞLIK AYARLAMA: 
    # Senin dosyanın orijinal başlıkları 3. satırda (indeks 2). Onları çekiyoruz.
    orijinal_basliklar = df_raw.iloc[2].values
    df_range.columns = orijinal_basliklar

    # 4. SÜTUN SEÇİMİ: 'Üst Birim'den başla ve sağa doğru 17 sütun git
    if 'Üst Birim' in df_range.columns:
        ust_birim_idx = list(df_range.columns).index('Üst Birim')
        final_df = df_range.iloc[:, ust_birim_idx : ust_birim_idx + 17]
        
        # Sütun isimlerindeki boşlukları temizle
        final_df.columns = [str(col).strip() for col in final_df.columns]

        # 5. BİÇİMLENDİRME: Oranları yüzdeye çevir
        oran_sutunlari = [col for col in final_df.columns if 'Oran' in col or 'Hedef' in col]
        
        styled_df = final_df.style.format({col: "{:.1%}" for col in oran_sutunlari})\
            .set_properties(**{
                'text-align': 'center',
                'font-size': '12px',
                'white-space': 'nowrap', # Sütunları en dar hale getirir
                'border': '1px solid #eeeeee'
            })\
            .hide(axis="index") # Sol taraftaki gereksiz satır numaralarını gizle

        st.write("### Belirlenen Aralık Önizlemesi")
        st.write(styled_df)

        # 6. FOTOĞRAF ÇIKTISI
        if st.button("🖼️ Fotoğrafı Hazırla ve İndir"):
            with st.spinner('Resim oluşturuluyor...'):
                resim_yolu = "ozel_aralik_cikti.png"
                dfi.export(styled_df, resim_yolu)
                
                with open(resim_yolu, "rb") as file:
                    st.download_button(
                        label="Görseli Kaydet",
                        data=file,
                        file_name="satir_26_43_rapor.png",
                        mime="image/png"
                    )
    else:
        st.error("Sütun başlıkları bulunamadı. Lütfen doğru Excel formatını yüklediğinizden emin olun.")
