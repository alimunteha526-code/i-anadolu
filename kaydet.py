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

            # 3. VERİ TEMİZLEME (Hata Engelleyici)
            oran_cols = [c for c in final_df.columns if 'Oran' in str(c) or 'Hedef' in str(c)]
            for col in oran_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            # 4. SIRALAMA (Büyükten Küçüğe)
            if target_col in final_df.columns:
                final_df = final_df.sort_values(by=target_col, ascending=False)

            # 5. BAŞLIK SATIRI OLUŞTURMA (İlk iki hücreyi kapsayan metin)
            baslik_satiri = pd.DataFrame(columns=final_df.columns)
            baslik_satiri.loc[0] = [""] * len(final_df.columns)
            # Metni ilk iki hücreye paylaştırıyoruz (veya yan yana görünmesini sağlıyoruz)
            baslik_satiri.iloc[0, 0] = "İÇ ANADOLU"
            baslik_satiri.iloc[0, 1] = "BÖLGESİ"
            
            final_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            # 6. GÜVENLİ FORMATLAMA FONKSİYONU
            def format_yuzde(x):
                try:
                    if pd.isna(x) or isinstance(x, str): return "-"
                    return "{:.1%}".format(float(x))
                except: return str(x)

            # 7. GÖRSEL STİL
            def stil_uygula(row):
                # Başlık satırı kontrolü (İlk hücrede İÇ ANADOLU yazıyorsa)
                if row.iloc[0] == "İÇ ANADOLU":
                    return ['background-color: #2c3e50; color: white; font-weight: bold; text-align: center; font-size: 14px'] * len(row)
                
                styles = [''] * len(row)
                if target_col in row.index:
                    val = row[target_col]
                    # Sayısal değerleri kontrol et (%5.8 sınırı)
                    try:
                        num_val = float(val)
                        if num_val > 0.058:
                            idx = row.index.get_loc(target_col)
                            styles[idx] = 'background-color: #e74c3c; color: white; font-weight: bold'
                    except: pass
                return styles

            # Formatı uygula
            for col in oran_cols:
                final_df[col] = final_df[col].apply(format_yuzde)

            # Tabloyu oluştur
            styled_df = final_df.style.apply(stil_uygula, axis=1)\
                .set_properties(**{
                    'text-align': 'center',
                    'font-size': '12px',
                    'border': '1px solid black',
                    'white-space': 'nowrap'
                })\
                .hide(axis="index")

            st.write("### Düzenlenmiş ve Sıralanmış Liste")
            st.write(styled_df)

            if st.button("🖼️ Fotoğrafı Hazırla"):
                resim_adi = "zayi_listesi.png"
                dfi.export(styled_df, resim_adi, table_conversion='chrome')
                with open(resim_adi, "rb") as file:
                    st.download_button("Görseli Kaydet", file, "zayi_listesi.png", "image/png")
        else:
            st.error("'Üst Birim' sütunu bulunamadı!")
                
    except Exception as e:
        st.error(f"Sistemsel Hata: {e}")
