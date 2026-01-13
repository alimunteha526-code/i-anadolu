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
    /* Sıfırla butonu için kırmızı stil */
    div[data-testid="stColumn"]:nth-child(2) .stButton>button {
        background-color: #d32f2f !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1>👓 ATASUN OPTİK</h1>", unsafe_allow_html=True)

# --- SESSION STATE (BELLEK) ---
if 'db' not in st.session_state:
    st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
if 'okutulanlar' not in st.session_state:
    st.session_state.okutulanlar = set()

# --- 1. ADIM: DOSYA YÜKLEME ---
with st.expander("📁 Sipariş Listesi Yükle (Excel veya ODS)", expanded=st.session_state.db.empty):
    # MIME tipi hatalarını önlemek için type=None yapıldı
    yuklenen_dosya = st.file_uploader("Dosyayı seçin veya buraya sürükleyin", type=None)
    
    if yuklenen_dosya:
        try:
            # Uzantıya göre motor seçimi
            if yuklenen_dosya.name.lower().endswith('.ods'):
                df_temp = pd.read_excel(yuklenen_dosya, engine='odf')
            else:
                df_temp = pd.read_excel(yuklenen_dosya)
            
            st.info("Sütunları eşleştirin:")
            c1, c2, c3 = st.columns(3)
            s_no_col = c1.selectbox("Sipariş No", df_temp.columns)
            s_isim_col = c2.selectbox("Müşteri Adı", df_temp.columns)
            s_pers_col = c3.selectbox("Personel No", df_temp.columns)
            
            if st.button("Listeye Ekle / Güncelle"):
                yeni_veri = df_temp[[s_no_col, s_isim_col, s_pers_col]].copy()
                yeni_veri.columns = ['Sipariş No', 'Müşteri Adı', 'Personel No']
                
                # Veri formatlama
                yeni_veri['Sipariş No'] = yeni_veri['Sipariş No'].astype(str).str.strip().str.upper()
                yeni_veri['Personel No'] = pd.to_numeric(yeni_veri['Personel No'], errors='coerce').fillna(0).astype(int).astype(str)
                
                # Birleştirme ve Mükerrer Kontrolü (Sipariş No bazlı)
                st.session_state.db = pd.concat([st.session_state.db, yeni_veri]).drop_duplicates(subset=['Sipariş No'], keep='last')
                st.success(f"✅ Liste güncellendi. Toplam Kayıt: {len(st.session_state.db)}")
        except Exception as e:
            st.error(f"Dosya okuma hatası: {e}")

st.divider()

# --- 2. ADIM: BARKOD OKUTMA ---
if not st.session_state.db.empty:
    with st.form(key='barkod_form', clear_on_submit=True):
        st.markdown("### 📲 Barkodu Okutun")
        input_kod = st.text_input("", placeholder="Barkodu okutun...").strip().upper()
        submit = st.form_submit_button("SORGULA")

    if submit and input_kod:
        match = st.session_state.db[st.session_state.db['Sipariş No'] == input_kod]
        if not match.empty:
            isim = match['Müşteri Adı'].iloc[0]
            st.success(f"✅ DOĞRU: {isim}")
            st.session_state.okutulanlar.add(input_kod)
        else:
            st.error(f"❌ LİSTEDE YOK: {input_kod}")

# --- 3. ADIM: RAPORLAMA VE PANELİ SIFIRLA ---
st.divider()
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    btn_rapor = st.button("📊 Eksikleri Listele")

with col_btn2:
    if st.button("🔄 Paneli Sıfırla"):
        st.session_state.db = pd.DataFrame(columns=['Sipariş No', 'Müşteri Adı', 'Personel No'])
        st.session_state.okutulanlar = set()
        st.rerun()

if btn_rapor:
    eksik_df = st.session_state.db[~st.session_state.db['Sipariş No'].isin(st.session_state.okutulanlar)].copy()
    
    if not eksik_df.empty:
        eksik_df.insert(0, 'Sıra No', range(1, len(eksik_df) + 1))
        st.markdown("### 📋 EKSİK SİPARİŞ LİSTESİ")
        st.dataframe(eksik_df, use_container_width=True, hide_index=True)
        
        st.markdown("#### 📥 İndirme Seçenekleri")
        d_col1, d_col2 = st.columns(2)

        # PDF OLARAK İNDİR (Hata Giderilmiş Bölüm)
        with d_col1:
            try:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", 'B', 14)
                pdf.cell(190, 10, "EKSIK SIPARIS LISTESI", ln=True, align='C')
                pdf.set_font("Arial", size=10)
                pdf.ln(5)
                # Tablo Başlığı
                pdf.cell(15, 8, "Sira", 1); pdf.cell(45, 8, "Siparis No", 1); pdf.cell(90, 8, "Musteri Adi", 1); pdf.cell(30, 8, "Pers. No", 1); pdf.ln()
                # Veriler
                for _, r in eksik_df.iterrows():
                    # Türkçe karakterleri temizleme (Latin-1 uyumu için)
                    isim_temiz = str(r['Müşteri Adı']).translate(str.maketrans("İıĞğÜüŞşÖöÇç", "IiGgUuSsOoCc"))
                    pdf.cell(15, 8, str(r['Sıra No']), 1)
                    pdf.cell(45, 8, str(r['Sipariş No']), 1)
                    pdf.cell(90, 8, isim_temiz[:40], 1)
                    pdf.cell(30, 8, str(r['Personel No']), 1)
                    pdf.ln()
                
                # Hata veren kısım düzeltildi:
                pdf_output = pdf.output()
                # Eğer çıktı bytes değilse (eski sürüm fpdf), bytes'a çevir:
                if isinstance(pdf_output, str):
                    pdf_output = pdf_output.encode('latin-1', 'replace')
                
                st.download_button("📄 PDF İndir", data=pdf_output, file_name="eksik_siparisler.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"PDF oluşturulurken hata: {e}")

        # ODS OLARAK İNDİR
        with d_col2:
            try:
                ods_buffer = io.BytesIO()
                with pd.ExcelWriter(ods_buffer, engine='odf') as writer:
                    eksik_df.to_excel(writer, index=False, sheet_name='Eksikler')
                st.download_button("📂 ODS İndir", data=ods_buffer.getvalue(), file_name="eksik_siparisler.ods", mime="application/vnd.oasis.opendocument.spreadsheet")
            except Exception as e:
                st.error(f"ODS oluşturulurken hata: {e}")
    else:
        st.success("Tüm siparişler tamamlanmış!")

# İstatistik alt bilgi
if not st.session_state.db.empty:
    st.markdown(f"<p style='text-align:center; color:#888;'>Sistemde {len(st.session_state.db)} kayıt var | {len(st.session_state.okutulanlar)} okutuldu.</p>", unsafe_allow_html=True)
