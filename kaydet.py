import streamlit as st
import pandas as pd
import io
from fpdf import FPDF

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(page_title="Atasun Optik - Takip Paneli", layout="centered")

# --- ATASUN KURUMSAL TASARIM (CSS) ---
st.markdown("""
    <style>
    .stApp { background-color: #FF671B; }
    .block-container {
        background-color: white;
        padding: 3rem;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
        margin-top: 2rem;
    }
    h1 { color: #333333; font-family: 'Arial Black', sans-serif; text-align: center; }
    .stButton>button { width: 100%; background-color: #333333 !important; color: white !important; font-weight: bold; border-radius: 10px !important; height: 3.5em; }
    .stDownloadButton>button { background-color: #007bff !important; color: white !important; border-radius: 10px !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1>👓 ATASUN OPTİK</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; font-weight:bold; color:#666;'>Açık Kapora Takip Paneli (Excel & ODS Destekli)</p>", unsafe_allow_html=True)

# Session State Hazırlığı
if 'db' not in st.session_state:
    st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
if 'okutulanlar' not in st.session_state:
    st.session_state.okutulanlar = set()

# --- 1. ADIM: DOSYA YÜKLEME VE BİRLEŞTİRME ---
with st.expander("📁 Sipariş Listesi Yükle (Excel veya ODS)", expanded=st.session_state.db.empty):
    yuklenen_dosya = st.file_uploader("Dosya seçiniz", type=['xlsx', 'ods'])
    
    if yuklenen_dosya:
        # Uzantıya göre okuma yapar (odfpy gereklidir)
        df_temp = pd.read_excel(yuklenen_dosya)
        
        st.info("Lütfen sütun eşleştirmelerini kontrol edin:")
        c1, c2, c3 = st.columns(3)
        s_no_col = c1.selectbox("Sipariş No Sütunu", df_temp.columns)
        s_isim_col = c2.selectbox("Müşteri Adı Sütunu", df_temp.columns)
        s_pers_col = c3.selectbox("Personel No Sütunu", df_temp.columns)
        
        if st.button("Listeye Ekle / Güncelle"):
            yeni_veri = df_temp[[s_no_col, s_isim_col, s_pers_col]].copy()
            yeni_veri.columns = ['Sipariş No', 'Müşteri Adı', 'Personel No']
            
            # Veri Temizleme
            yeni_veri['Sipariş No'] = yeni_veri['Sipariş No'].astype(str).str.strip().str.upper()
            yeni_veri['Personel No'] = pd.to_numeric(yeni_veri['Personel No'], errors='coerce').fillna(0).astype(int).astype(str)
            
            # Mevcut liste ile birleştir ve mükerrer kayıtları (Sipariş No bazlı) sil
            birlesik_df = pd.concat([st.session_state.db, yeni_veri]).drop_duplicates(subset=['Sipariş No'], keep='last')
            st.session_state.db = birlesik_df
            st.success(f"✅ Liste güncellendi. Toplam Kayıt: {len(st.session_state.db)}")

st.divider()

# --- 2. ADIM: BARKOD OKUTMA ---
if not st.session_state.db.empty:
    with st.form(key='barkod_form', clear_on_submit=True):
        st.markdown("### 📲 Barkodu Okutun")
        input_kod = st.text_input("", placeholder="Barkodu buraya okutun...").strip().upper()
        submit = st.form_submit_button("SORGULA")

    if submit and input_kod:
        match = st.session_state.db[st.session_state.db['Sipariş No'] == input_kod]
        if not match.empty:
            isim = match['Müşteri Adı'].iloc[0]
            st.success(f"✅ DOĞRU: {isim}")
            st.session_state.okutulanlar.add(input_kod)
        else:
            st.error(f"❌ LİSTEDE YOK: {input_kod}")

# --- İstatistik ---
if not st.session_state.db.empty:
    toplam = len(st.session_state.db)
    okunan = len(st.session_state.okutulanlar)
    st.info(f"İlerleme: {okunan} / {toplam} (Kalan: {toplam - okunan})")

# --- 3. ADIM: RAPORLAMA ---
st.divider()
if st.button("📊 Eksikleri Listele"):
    eksik_df = st.session_state.db[~st.session_state.db['Sipariş No'].isin(st.session_state.okutulanlar)].copy()
   
    if not eksik_df.empty:
        eksik_df.insert(0, 'Sıra No', range(1, len(eksik_df) + 1))
        st.markdown("### 📋 EKSİK SİPARİŞ LİSTESİ")
        st.dataframe(eksik_df, use_container_width=True, hide_index=True)
       
        col_pdf, col_csv = st.columns(2)
       
        with col_pdf:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(190, 10, "EKSIK SIPARIS LISTESI", ln=True, align='C')
            pdf.set_font("Arial", size=10)
            pdf.ln(5)
            # Başlıklar
            pdf.cell(15, 8, "Sira", 1)
            pdf.cell(45, 8, "Siparis No", 1)
            pdf.cell(90, 8, "Musteri Adi", 1)
            pdf.cell(30, 8, "Pers. No", 1)
            pdf.ln()
            # Satırlar
            for i, r in eksik_df.iterrows():
                # Türkçe karakter koruması
                isim_pdf = str(r['Müşteri Adı']).replace('İ','I').replace('ğ','g').replace('ü','u').replace('ş','s').replace('ö','o').replace('ç','c').replace('Ğ','G').replace('Ü','U').replace('Ş','S').replace('Ö','O').replace('Ç','C')
                pdf.cell(15, 8, str(r['Sıra No']), 1)
                pdf.cell(45, 8, str(r['Sipariş No']), 1)
                pdf.cell(90, 8, isim_pdf[:40], 1)
                pdf.cell(30, 8, str(r['Personel No']), 1)
                pdf.ln()
           
            pdf_bytes = pdf.output(dest='S').encode('latin-1', 'replace')
            st.download_button("📄 PDF İndir", data=pdf_bytes, file_name="Eksik_Siparisler.pdf", mime="application/pdf")

        with col_csv:
            csv_data = eksik_df.to_csv(index=False, encoding='utf-8-sig', sep=';')
            st.download_button("📂 CSV (Excel için) İndir", data=csv_data, file_name="Eksik_Siparisler.csv", mime="text/csv")
           
    else:
        st.balloons()
        st.success("Harika! Eksik siparişiniz bulunmuyor.")

if st.button("🔄 Tüm Sistemi Sıfırla"):
    st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
    st.session_state.okutulanlar = set()
    st.rerun()
