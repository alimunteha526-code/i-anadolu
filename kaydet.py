import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="İç Anadolu Raporu", layout="wide")

st.title("📊 Bölge Bazlı Rapor Oluşturucu")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    # 1. Ham veriyi oku
    df_raw = pd.read_excel(yuklenen_dosya, header=None)

    # 2. 26-43 aralığını al (İndeks 25:43)
    df_range = df_raw.iloc[25:43].copy()

    # 3. Başlıkları ata (Orijinal başlıklar 3. satırda / İndeks 2)
    orijinal_basliklar = df_raw.iloc[2].values
    df_range.columns = orijinal_basliklar

    # 4. Üst Birim'den itibaren 17 sütun al
    if 'Üst Birim' in df_range.columns:
        ust_birim_idx = list(df_range.columns).index('Üst Birim')
        final_df = df_range.iloc[:, ust_birim_idx : ust_birim_idx + 17].copy()
        final_df.columns = [str(col).strip() for col in final_df.columns]

        # --- ÖZEL BAŞLIK EKLEME VE STİLLENDİRME ---
        oran_sutunlari = [col for col in final_df.columns if 'Oran' in col or 'Hedef' in col]
        
        # Stil ayarları
        styled_df = final_df.style.format({col: "{:.1%}" for col in oran_sutunlari})\
            .set_table_styles([
                # Burası tablonun en üstüne birleştirilmiş başlık ekler
                {'selector': 'thead', 'props': [('display', 'table-header-group')]},
                {'selector': 'caption', 'props': [
                    ('caption-side', 'top'), 
                    ('color', 'white'), 
                    ('font-size', '16px'), 
                    ('font-weight', 'bold'),
                    ('text-align', 'center'),
                    ('background-color', '#2c3e50'), # Lacivert arka plan
                    ('padding', '10px')
                ]}
            ])\
            .set_properties(**{
                'text-align': 'center',
                'font-size': '12px',
                'white-space': 'nowrap',
                'border': '1px solid #ddd'
            })\
            .hide(axis="index") # Satır numaralarını gizle
            
        # Tablo başlığını ayarla
        styled_df.set_caption("İÇ ANADOLU BÖLGESİ")

        st.write("### Tablo Önizlemesi")
        st.write(styled_df)

        if st.button("🖼️ Fotoğraf Olarak İndir"):
            with st.spinner('Fotoğraf hazırlanıyor...'):
                resim_yolu = "ic_anadolu_raporu.png"
                # Başlık ile birlikte dışa aktar
                dfi.export(styled_df, resim_yolu)
                
                with open(resim_yolu, "rb") as file:
                    st.download_button("Görseli Kaydet", file, "rapor.png", "image/png")
    else:
        st.error("'Üst Birim' sütunu bulunamadı!")
