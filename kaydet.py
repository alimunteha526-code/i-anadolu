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
    /* Sıfırla butonu için özel kırmızı stil */
    div[data-testid="stColumn"]:nth-child(2) .stButton>button {
        background-color: #d32f2f !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1>👓 ATASUN OPTİK</h1>", unsafe_allow_html=True)

# --- SESSION STATE (BELLEK) YÖNETİMİ ---
if 'db' not in st.session_state:
    st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
if 'okutulanlar' not in st.session_state:
    st.session_state.okutulanlar = set()

# --- 1. ADIM: DOSYA YÜKLEME ---
with st.expander("📁 Sipariş Listesi Yükle (Excel veya ODS)", expanded=st.session_state.db.empty):
    # MIME tipi kısıtlamasını aşmak için type=None yapıldı
    yuklenen_dosya = st.file_uploader("Dosyayı seçin veya sürükleyin", type=None)
    
    if yuklenen_dosya:
        try:
            # Uzantıya göre okuma motoru seçimi
            if yuklenen_dosya.name.lower().endswith('.ods'):
                df_temp = pd.read_excel(yuklenen_dosya, engine='odf')
            else:
                df_temp = pd.read_excel(yuklenen_dosya)
            
            st.info("Sütun eşleştirmelerini kontrol edin:")
            c1, c2, c3 = st.columns(3)
            s_no_col = c1.selectbox("Sipariş No", df_temp.columns)
            s_isim_col = c2.selectbox("Müşteri Adı", df_temp.columns)
            s_pers_col = c3.selectbox("Personel No", df_temp.columns)
            
            if st.button("Listeye Ekle / Güncelle"):
                yeni_veri = df_temp[[s_no_col, s_isim_col, s_pers_col]].copy()
                yeni_veri.columns = ['Sipariş No', 'Müşteri Adı', 'Personel No']
                
                # Temizlik
                yeni_veri['Sipariş No'] = yeni_veri['Sipariş No'].astype(str).str.strip().str.upper()
                yeni_veri['Personel No'] = pd.to_numeric(yeni_veri['Personel No'], errors='coerce').fillna(0).astype(int).astype(str)
                
                # Mevcut veriye ekleme ve tekrarları silme
                st.session_state.db = pd.concat([st.session_state.db, yeni_veri]).drop_duplicates(subset=['Sipariş No'], keep='last')
                st.success(f"✅ Başarılı! Toplam kayıt sayısı: {len(st.session_state.db)}")
        except Exception as e:
            st.error(f"Dosya yüklenirken bir hata oluştu: {e}")

st.divider()

# --- 2. ADIM: BARKOD OKUTMA ---
if not st.session_state.db.empty:
    with st.form(key='barkod_form', clear_on_submit=True):
        st.markdown("### 📲 Barkod Okutma")
        input_kod = st.text_input("Barkodu okutun ve Sorgula'ya basın", placeholder="...").strip().upper()
        submit = st.form_submit_button("SORGULA")

    if submit and input_kod:
        match = st.session_state.db[st.session_state.db['Sipariş No'] == input_kod]
        if not match.empty:
            isim = match['Müşteri Adı'].iloc[0]
            st.success(f"✅ EŞLEŞTİ: {isim}")
            st.session_state.okutulanlar.add(input_kod)
        else:
            st.error(f"❌ KAYIT BULUNAMADI: {input_kod}")

# --- 3. ADIM: RAPORLAMA VE PANELİ SIFIRLA ---
st.divider()
col_aksiyon1, col_aksiyon2 = st.columns(2)

with col_aksiyon1:
    btn_eksik = st.button("📊 Eksikleri Listele")

with col_aksiyon2:
    if st.button("🔄 Paneli Sıfırla"):
        st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
        st.session_state.okutulanlar = set()
        st.rerun()

# Eksik listesi tetiklendiğinde
if btn_eksik:
    eksik_df = st.session_state.db[~st.session_state.db['Sipariş No'].isin(st.session_state.okutulanlar)].copy()
    
    if not eksik_df.empty:
        eksik_df.insert(0, 'Sıra No', range(1, len(eksik_df) + 1))
        st.markdown("### 📋 EKSİK SİPARİŞ LİSTESİ")
        st.dataframe(eksik_df, use_container_width=True, hide_index=True)
        
        st.markdown("#### 📥 Farklı Formatta İndir")
        d_col1, d_col2 = st.columns(2)

        # PDF İndirme (Hata Giderilmiş Versiyon)
        with d_col1:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(190, 10, "EKSIK SIPARIS LISTESI", ln=True, align='C')
            pdf.set_font("Arial", size=10)
            pdf.ln(5)
            # Tablo Başlıkları
            pdf.cell(15, 8, "Sira", 1); pdf.cell(45, 8, "Siparis No", 1); pdf.cell(90, 8, "Musteri Adi", 1); pdf.cell(30, 8, "Pers. No", 1); pdf.ln()
            # Veriler
            for _, r in eksik_df.iterrows():
                # Latin-1 uyumluluğu için Türkçe karakter değişimi
                ad_pdf = str(r['Müşteri Adı']).translate(str.maketrans("İıĞğÜüŞşÖöÇç", "IiGgUuSsOoCc"))
                pdf.cell(15, 8, str(r['Sıra No']), 1)
                pdf.cell(45, 8, str(r['Sipariş No']), 1)
                pdf.cell(90, 8, ad_pdf[:40], 1)
                pdf.cell(30, 8, str(r['Personel No']), 1)
                pdf.ln()
            
            # FPDF2 için en güvenli bayt çıktısı
            pdf_bytes = pdf.output()
            if isinstance(pdf_bytes, str): # Eğer eski sürüm fpdf ise
                pdf_bytes = pdf_bytes.encode('latin-1', 'replace')
                
            st.download_button("📄 PDF Olarak İndir", data=pdf_bytes, file_name="eksik_liste.pdf", mime="application/pdf")

        # ODS İndirme
        with d_col2:
            ods_buffer = io.BytesIO()
            with pd.ExcelWriter(ods_buffer, engine='odf') as writer:
                eksik_df.to_excel(writer, index=False, sheet_name='Eksik_Siparisler')
            
            st.download_button("📂 ODS Olarak İndir", data=ods_buffer.getvalue(), file_name="eksik_liste.ods", mime="application/vnd.oasis.opendocument.spreadsheet")
            
    else:
        st.success("Harika! Tüm siparişler okutulmuş, eksik liste boş.")

# Alt Bilgi
if not st.session_state.db.empty:
    st.markdown(f"<p style='text-align:center; color:gray;'>Sistemde {len(st.session_state.db)} kayıt var | {len(st.session_state.okutulanlar)} adet okutuldu.</p>", unsafe_allow_html=True)
