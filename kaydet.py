import streamlit as st
import pandas as pd
import io
from fpdf import FPDF

st.set_page_config(page_title="Zayi Düzenleme Paneli", layout="wide")
st.title("📋 İÇ ANADOLU AEL ZAYİ LİSTESİ DÜZENLEME PANELİ")

# --- TÜRKÇE KARAKTER TEMİZLEYİCİ FONKSİYON ---
def tr_to_en(text):
    if not isinstance(text, str):
        return text
    mapping = {
        "İ": "I", "ı": "i", "Ş": "S", "ş": "s", "Ğ": "G", "ğ": "g",
        "Ç": "C", "ç": "c", "Ö": "O", "ö": "o", "Ü": "U", "ü": "u"
    }
    for tr_char, en_char in mapping.items():
        text = text.replace(tr_char, en_char)
    return text

yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    try:
        df_full = pd.read_excel(yuklenen_dosya, header=2)
        df_full.columns = [str(c).strip() for c in df_full.columns]

        target_sort_col = 'Net Satış Miktarı (Cam)'
        zayi_oran_col = 'Toplam Cam Zayi Oranı'

        if 'Üst Birim' in df_full.columns:
            start_col = df_full.columns.get_loc('Üst Birim')
            final_df = df_full.iloc[22:40, start_col : start_col + 17].copy()

            numeric_cols = [c for c in final_df.columns if any(x in str(c) for x in ['Oran', 'Hedef', 'Miktarı', 'Adet'])]
            for col in numeric_cols:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            if target_sort_col in final_df.columns:
                final_df = final_df.sort_values(by=target_sort_col, ascending=False)

            baslik_satiri = pd.DataFrame(columns=final_df.columns)
            baslik_satiri.loc[0] = [""] * len(final_df.columns)
            baslik_satiri.iloc[0, 0] = "IC ANADOLU BOLGESI" # PDF için temizlendi
            
            report_df = pd.concat([baslik_satiri, final_df], ignore_index=True)

            st.write("### Liste Önizlemesi")
            st.dataframe(report_df)

            col1, col2 = st.columns(2)

            with col1:
                # --- EXCEL ÇIKTISI (Hatasız) ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                    report_df.to_excel(writer, index=False, sheet_name='Zayi_Raporu')
                    workbook = writer.book
                    worksheet = writer.sheets['Zayi_Raporu']
                    # (Buradaki Excel biçimlendirme kodlarını önceki mesajdaki gibi kullanabilirsiniz)
                st.download_button("📥 Renkli Excel İndir", output.getvalue(), "ic_anadolu_renkli.xlsx")

            with col2:
                # --- GÜVENLİ PDF ÇIKTISI ---
                try:
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 10)
                    pdf.cell(0, 10, "IC ANADOLU BOLGESI RAPORU", ln=True, align='C')
                    pdf.set_font("Helvetica", size=6)
                    
                    col_width = pdf.epw / len(report_df.columns)
                    
                    for i, row in report_df.iterrows():
                        is_header = "IC ANADOLU" in str(row.iloc[0])
                        for col_idx, item in enumerate(row):
                            col_name = report_df.columns[col_idx]
                            
                            # PDF için metni temizle (Karakter hatasını önler)
                            val = tr_to_en(str(item)) if pd.notna(item) else ""
                            
                            if is_header:
                                pdf.set_fill_color(31, 78, 120)
                                pdf.set_text_color(255, 255, 255)
                            elif col_name == zayi_oran_col and isinstance(item, (float, int)) and item > 0.058:
                                pdf.set_fill_color(255, 0, 0)
                                pdf.set_text_color(255, 255, 255)
                            else:
                                pdf.set_fill_color(255, 255, 255)
                                pdf.set_text_color(0, 0, 0)
                            
                            if isinstance(item, float) and 0 < item < 1:
                                val = f"{item:.1%}"
                            
                            pdf.cell(col_width, 7, val[:12], border=1, fill=True, align='C')
                        pdf.ln()

                    st.download_button("📥 Renkli PDF İndir", bytes(pdf.output()), "ic_anadolu_renkli.pdf")
                except Exception as e:
                    st.error(f"PDF Karakter Hatası Giderilemedi: {e}")

    except Exception as e:
        st.error(f"Sistem Hatası: {e}")
