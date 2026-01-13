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
    # 'type' kısıtlamasını kaldırarak MIME tipi hatalarını bypass ediyoruz
    yuklenen_dosya = st.file_uploader("Dosya seçiniz (.xlsx veya .ods)", type=None)
    
    if yuklenen_dosya:
        try:
            # Dosya uzantısına göre uygun motoru seçiyoruz
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
                
                # Veri Temizleme
                yeni_veri['Sipariş No'] = yeni_veri['Sipariş No'].astype(str).str.strip().str.upper()
                yeni_veri['Personel No'] = pd.to_numeric(yeni_veri['Personel No'], errors='coerce').fillna(0).astype(int).astype(str)
                
                # Birleştirme ve Mükerrer Kontrolü
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

# --- 3. ADIM: RAPORLAMA ---
if st.button("📊 Eksikleri Listele"):
    eksik_df = st.session_state.db[~st.session_state.db['Sipariş No'].isin(st.session_state.okutulanlar)].copy()
    if not eksik_df.empty:
        st.dataframe(eksik_df, use_container_width=True)
        
        # CSV Çıktısı (UTF-8 SIG ile Türkçe karakter desteği)
        csv_data = eksik_df.to_csv(index=False, encoding='utf-8-sig', sep=';')
        st.download_button("📂 Eksik Listesini İndir", data=csv_data, file_name="eksikler.csv")
    else:
        st.success("Tüm siparişler tamam!")

if st.sidebar.button("🔄 Sistemi Sıfırla"):
    st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
    st.session_state.okutulanlar = set()
    st.rerun()
