import streamlit as st
import pandas as pd
import dataframe_image as dfi
import os

# Sayfa ayarları
st.set_page_config(page_title="Rapor Oluşturucu", layout="wide")

st.title("📊 Profesyonel Cam Zayi Raporu")

# Gerekli fonksiyon: Renklendirme
def kirmizi_kutucuk(val):
    if isinstance(val, (int, float)) and val > 0.058:
        return 'background-color: #e74c3c; color: white; font-weight: bold'
    return ''

yuklenen_dosya = st.file_uploader("Excel Dosyanızı Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        # 1. DOSYAYI OKU: İlk 2 satırı atla, 3. satırı başlık yap
        df = pd.read_excel(yuklenen_dosya, skiprows=2)
        
        # Sütun isimlerindeki boşlukları temizle (Hata almamak için kritik)
        df.columns = [str(c).strip() for c in df.columns]

        # 2. SATIR VE SÜTUN AYIKLAMA
        # Excel 26. satır -> Python df içinde 23. satıra denk gelir
        # Üst Birim sütunundan itibaren 17 sütun al
        if 'Üst Birim' in df.columns:
            baslangic_idx = df.columns.get_loc('Üst Birim')
            
            # 26-39. satırlar (Üst kısım)
            ust_df = df.iloc[23:37, baslangic_idx : baslangic_idx + 17].copy()
            
            # 41-43. satırlar (Alt kısım)
            alt_df = df.iloc[38:41, baslangic_idx : baslangic_idx + 17].copy()

            # 3. ARA BAŞLIK SATIRI (40. satır yerine)
            ara_baslik = pd.DataFrame(columns=ust_df.columns)
            ara_baslik.loc[0] = [""] * len(ust_df.columns)
            ara_baslik.iloc[0, 0] = "İÇ ANADOLU BÖLGESİ" # İlk hücreye yaz

            # Hepsini birleştir
            final_df = pd.concat([ust_df, ara_baslik, alt_df], ignore_index=True)

            # 4. GÖRSEL STİL VE BİÇİMLENDİRME
            oran_cols = [c for c in final_df.columns if 'Oran' in c or 'Hedef' in c]
            
            # Tabloyu özelleştir
            styled_df = final_df.style.format({c: "{:.1%}" for c in oran_cols}, na_rep="")
            
            # %5.8 üzeri kırmızı yap (Sütun adı tam eşleşmeli)
            target_col = 'Toplam Cam Zayi Oranı'
            if target_col in final_df.columns:
                styled_df = styled_df.applymap(kirmizi_kutucuk, subset=[target_col])

            # Ara başlığı renklendir (Satır bazlı kontrol)
            def satir_stili(row):
                if row.iloc[0] == "İÇ ANADOLU BÖLGESİ":
                    return ['background-color: #2c3e50; color: white; font-weight: bold'] * len(row)
                return [''] * len(row)
            
            styled_df = styled_df.apply(satir_stili, axis=1)\
                .set_properties(**{
                    'text-align': 'center',
                    'font-size': '12px',
                    'border': '1px solid #ddd',
                    'white-space': 'nowrap'
                })\
                .hide(axis="index")

            st.write("### Tablo Önizlemesi")
            st.write(styled_df)

            # 5. FOTOĞRAF OLARAK İNDİR
            if st.button("🖼️ Fotoğraf Olarak İndir"):
                with st.spinner('Resim hazırlanıyor...'):
                    # Geçici dosya adı
                    resim_yolu = "cikti_rapor.png"
                    dfi.export(styled_df, resim_yolu)
                    
                    with open(resim_yolu, "rb") as file:
                        st.download_button("Resmi Kaydet", file, "rapor.png", "image/png")
        else:
            st.error("Hata: 'Üst Birim' sütunu bulunamadı. Lütfen Excel başlıklarını kontrol edin.")

    except Exception as e:
        st.error(f"Sistemsel bir hata oluştu: {e}")
