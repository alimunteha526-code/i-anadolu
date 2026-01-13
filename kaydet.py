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
    .reset-button>button { background-color: #dc3545 !important; } /* Sıfırla butonu için kırmızı ton */
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1>👓 ATASUN OPTİK</h1>", unsafe_allow_html=True)

# Session State Hazırlığı
if 'db' not in st.session_state:
    st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
if 'okutulanlar' not in st.session_state:
    st.session_state.okutulanlar = set()

# --- 1. ADIM: DOSYA YÜKLEME ---
with st.expander("📁 Sipariş Listesi Yükle (Excel veya ODS)", expanded=st.session_state.db.empty):
    yuklenen_dosya = st.file_uploader("Dosya seçiniz", type=None)
    
    if yuklenen_dosya:
        try:
            if yuklenen_dosya.name.endswith('.ods'):
                df_temp = pd.read_excel(yuklenen_dosya, engine='odf')
            else:
                df_temp = pd.read_excel(yuklenen_dosya)
            
            st.info("Sütun eşleştirmelerini kontrol edin:")
            c1, c2, c3 = st.columns(3)
            s_no_col = c1.selectbox("Sipariş No", df_temp.columns)
            s_isim_col = c2.selectbox("Müşteri Adı", df_temp.columns)
            s_pers_col = c3.selectbox("Personel No", df_temp.columns)
            
            if st.button("Listeye Ekle"):
                yeni_veri = df_temp[[s_no_col, s_isim_col, s_pers_col]].copy()
                yeni_veri.columns = ['Sipariş No', 'Müşteri Adı', 'Personel No']
                yeni_veri['Sipariş No'] = yeni_veri['Sipariş No'].astype(str).str.strip().str.upper()
                yeni_veri['Personel No'] = pd.to_numeric(yeni_veri['Personel No'], errors='coerce').fillna(0).astype(int).astype(str)
                
                birlesik_df = pd.concat([st.session_state.db, yeni_veri]).drop_duplicates(subset=['Sipariş No'], keep='last')
                st.session_state.db = birlesik_df
                st.success(f"✅ Liste güncellendi. Toplam: {len(st.session_state.db)}")
        except Exception as e:
            st.error(f"Dosya okuma hatası: {e}")

st.divider()

# --- 2. ADIM: BARKOD OKUTMA ---
if not st.session_state.db.empty:
    with st.form(key='barkod_form', clear_on_submit=True):
        input_kod = st.text_input("📲 Barkodu Okutun").strip().upper()
        submit = st.form_submit_button("SORGULA")

    if submit and input_kod:
        match = st.session_state.db[st.session_state.db['Sipariş No'] == input_kod]
        if not match.empty:
            isim = match['Müşteri Adı'].iloc[0]
            st.success(f"✅ DOĞRU: {isim}")
            st.session_state.okutulanlar.add(input_kod)
        else:
            st.error(f"❌ LİSTEDE YOK: {input_kod}")

# --- 3. ADIM: RAPORLAMA VE SIFIRLAMA BUTONLARI ---
st.divider()
col_rapor, col_sifirla = st.columns(2)

with col_rapor:
    btn_eksik = st.button("📊 Eksikleri Listele")

with col_sifirla:
    if st.button("🔄 Paneli Sıfırla"):
        st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
        st.session_state.okutulanlar = set()
        st.rerun()

if btn_eksik:
    eksik_df = st.session_state.db[~st.session_state.db['Sipariş No'].isin(st.session_state.okutulanlar)].copy()
    
    if not eksik_df.empty:
        eksik_df.insert(0, 'Sıra No', range(1, len(eksik_df) + 1))
        st.markdown("### 📋 EKSİK SİPARİŞ LİSTESİ")
        st.dataframe(eksik_df, use_container_width=True, hide_index=True)
        
        st.markdown("#### 📥 Listeyi İndir")
        download_col1, download_col2 = st.columns(2)

        # --- PDF OLARAK İNDİR ---
        with download_col1:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(190, 10, "EKSIK SIPARIS LISTESI", ln=True, align='C')
            pdf.set_font("Arial", size=10)
            pdf.ln(5)
            pdf.cell(15, 8, "Sira", 1); pdf.cell(45, 8, "Siparis No", 1); pdf.cell(90, 8, "Musteri Adi", 1); pdf.cell(30, 8, "Pers. No", 1); pdf.ln()
            for i, r in eksik_df.iterrows():
                isim_temiz = str(r['Müşteri Adı']).replace('İ','I').replace('ğ','g').replace('ü','u').replace('ş','s').replace('ö','o').replace('ç','c').replace('Ğ','G').replace('Ü','U').replace('Ş','S').replace('Ö','O').replace('Ç','C')
                pdf.cell(15, 8, str(r['Sıra No']), 1)
                pdf.cell(45, 8, str(r['Sipariş No']), 1)
                pdf.cell(90, 8, isim_temiz[:40], 1)
                pdf.cell(30, 8, str(r['Personel No']), 1)
                pdf.ln()
            pdf_output = pdf.output(dest='S').encode('latin-1', 'replace')
            st.download_button("📄 PDF İndir", data=pdf_output, file_name="eksik_liste.pdf", mime="application/pdf")

        # --- ODS OLARAK İNDİR ---
        with download_col2:
            output_ods = io.BytesIO()
            with pd.ExcelWriter(output_ods, engine='odf') as writer:
                eksik_df.to_excel(writer, index=False, sheet_name='Eksikler')
            st.download_button("📂 ODS İndir", data=output_ods.getvalue(), file_name="eksik_liste.ods", mime="application/vnd.oasis.opendocument.spreadsheet")
            
    else:
        st.success("Tüm siparişler tamamlanmış, eksik yok!")
