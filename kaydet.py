import streamlit as st
import pandas as pd
import dataframe_image as dfi

# ... (Dosya yükleme ve veri ayıklama kısımları aynı kalıyor) ...

        # --- YÜZDE VE RENK BİÇİMLENDİRME ---
        oran_sutunlari = [col for col in final_df.columns if 'Oran' in col or 'Hedef' in col]
        
        # Kırmızı kutu fonksiyonu
        def kirmizi_kutu(val):
            # Değer 0.058'den (yani %5.8) büyükse arka planı kırmızı, yazıyı beyaz yap
            color = 'background-color: #e74c3c; color: white; font-weight: bold' if isinstance(val, (int, float)) and val > 0.058 else ''
            return color

        # Stili uygula
        styled_df = final_df.style.format({col: "{:.1%}" for col in oran_sutunlari})\
            .applymap(kirmizi_kutu, subset=['Toplam Cam Zayi Oranı'])\
            .set_table_styles([
                {'selector': 'caption', 'props': [
                    ('caption-side', 'top'), ('color', 'white'), ('font-size', '16px'), 
                    ('font-weight', 'bold'), ('text-align', 'center'),
                    ('background-color', '#2c3e50'), ('padding', '10px')
                ]}
            ])\
            .set_properties(**{
                'text-align': 'center',
                'font-size': '12px',
                'white-space': 'nowrap',
                'border': '1px solid #ddd'
            })\
            .hide(axis="index")
            
        styled_df.set_caption("İÇ ANADOLU BÖLGESİ - RİSKLİ MAĞAZA TAKİBİ")

        st.write("### %5.8 Üzeri Kırmızı ile İşaretlenmiştir")
        st.write(styled_df)
# ... (İndirme butonu kısmı aynı) ...
