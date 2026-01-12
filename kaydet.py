import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. Excel'i oku (Header 2. satırda - İndeks 2)
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            target_col = 'Toplam Cam Zayi Oranı'
            
            # 2. SATIR SEÇİMİ (Excel 26-43 aralığı)
            final_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            # 3. VERİ TEMİZLEME VE SIRALAMA (Kritik Bölüm)
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            
            # Oran sütunlarını sayıya çevir (Hata almamak ve doğru sıralamak için)
            for col in oran_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            # --- SIRALAMA İŞLEMİ (Büyükten Küçüğe) ---
            if target_col in final_df.columns:
                final_df = final_df.sort_values(by=target_col, ascending=False)

            # 4. EN BAŞA BÖLGE SATIRI EKLEME (Sıralamadan sonra ekliyoruz ki en üstte kalsın)
            baslik_satiri = pd.DataFrame(columns=final_df.columns)
            baslik_satiri.loc[0] = [""] * len(final_df.columns)
            baslik_satiri.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ"
            
            final_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            # 5. GÖRSEL STİL FONKSİYONLARI
            def stil_uygula(row):
                if row.iloc[0] == "İÇ ANADOLU BÖLGESİ":
                    return ['background-color: #2c3e50; color: white; font-weight: bold; text-align: center'] * len(row)
                
                styles = [''] * len(row)
                if target_col in row.index:
                    val = row[target_col]
                    # %5.8 (0.058) üzerindeyse kırmızı yap
                    if pd.notnull(val) and isinstance(val, (int, float)) and val > 0.058:
                        idx = row.index.get_loc(target_col)
                        styles[idx] = 'background-color: #e74c3c; color: white; font-weight: bold'
                return styles

            # 6. TABLO FORMATLAMA
            styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols}, na_rep="-")\
                .apply(stil_uygula, axis=1)\
                .set_properties(**{
                    'text-align': 'center',
                    'font-size': '12px',
                    'border': '1px solid black',
                    'white-space': 'nowrap'
                })\
                .hide(axis="index")

            st.write("### Sıralanmış Liste Önizlemesi (Büyükten Küçüğe)")
            st.write(styled_df)

            # 7. FOTOĞRAF ÇIKTISI
            if st.button("🖼️ Fotoğrafı Hazırla ve İndir"):
                with st.spinner('Görsel oluşturuluyor...'):
                    resim_adi = "zayi_sirali_liste.png"
                    dfi.export(styled_df, resim_adi, table_conversion='chrome')
                    
                    with open(resim_adi, "rb") as file:
                        st.download_button("Görseli Kaydet", file, "zayi_listesi.png", "image/png")
        else:
            st.error("'Üst Birim' sütunu bulunamadı!")
                
    except Exception as e:
        st.error(f"Hata detayı: {e}")
