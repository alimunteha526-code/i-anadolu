import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Rapor Oluşturucu", layout="wide")

st.title("📊 Satır Arası Başlıklı Cam Zayi Raporu")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Ham veriyi oku (3. satır başlık - İndeks 2)
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        
        # 2. Üst Birim'den itibaren 17 sütunu al
        ust_birim_idx = df_full.columns.get_loc('Üst Birim')
        df = df_full.iloc[:, ust_birim_idx : ust_birim_idx + 17].copy()
        df.columns = [str(c).strip() for c in df.columns]

        # 3. SATIR AYARLARI (Excel 26-43 aralığı)
        # Excel 26. satır -> df indeksi 23
        # Excel 40. satır -> df indeksi 37
        ust_kisim = df.iloc[23:37].copy() # 26'dan 39'a kadar olan mağazalar
        alt_kisim = df.iloc[38:41].copy() # 41'den 43'e kadar olan mağazalar

        # 4. ARA BAŞLIK SATIRI OLUŞTURMA (40. satır yerine)
        ara_baslik = pd.DataFrame(index=[37], columns=df.columns)
        ara_baslik.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ" # İlk hücreye yaz
        # Diğer hücreleri boş bırak (birleşmiş görünecek)
        ara_baslik.fillna("", inplace=True)

        # 5. Tabloları Birleştir
        final_df = pd.concat([ust_kisim, ara_baslik, alt_kisim])

        # --- GÖRSEL STİL ---
        def stil_uygula(row):
            styles = [''] * len(row)
            # Eğer satır bizim ara başlığımızsa (İç Anadolu yazıyorsa)
            if row['Üst Birim'] == "İÇ ANADOLU BÖLGESİ":
                return ['background-color: #2c3e50; color: white; font-weight: bold; text-align: center'] * len(row)
            
            # Normal satırlar için oran kontrolü (%5.8 üzeri kırmızı)
            val = row.get('Toplam Cam Zayi Oranı', 0)
            if isinstance(val, (int, float)) and val > 0.058:
                # Sadece o hücreyi kırmızı yap (indeksini bulmamız lazım)
                idx = list(row.index).index('Toplam Cam Zayi Oranı')
                styles[idx] = 'background-color: #e74c3c; color: white; font-weight: bold'
            
            return styles

        oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
        
        styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols if c in final_df.columns}, na_rep="")\
            .apply(stil_uygula, axis=1)\
            .set_properties(**{
                'text-align': 'center',
                'font-size': '12px',
                'border': '1px solid #ddd',
                'white-space': 'nowrap'
            })\
            .hide(axis="index")

        st.write("### Tablo Önizlemesi (40. Satır Başlık Yapıldı)")
        st.write(styled_df)

        if st.button("🖼️ Fotoğraf Olarak İndir"):
            resim = "ara_baslikli_rapor.png"
            dfi.export(styled_df, resim, table_conversion='chrome')
            with open(resim, "rb") as f:
                st.download_button("Resmi Kaydet", f, "rapor.png", "image/png")

    except Exception as e:
        st.error(f"Hata: {e}")
