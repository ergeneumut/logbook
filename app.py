import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
import os
import base64
import streamlit.components.v1 as components
from collections import defaultdict

# Sayfa Yapılandırması
st.set_page_config(page_title="Hızlı Logbook Otomasyonu", layout="wide")

# Görselleri Base64 formatına çeviren yardımcı fonksiyon
def get_base64_image(file_path):
    if os.path.exists(file_path):
        with open(file_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        ext = os.path.splitext(file_path)[1].lower()
        mime = "image/png" if ext == ".png" else "image/jpeg"
        return f"data:{mime};base64,{encoded_string}"
    return ""

# Görselleri klasörden oku
plane_base64 = get_base64_image("pngegg.jpg")
hangar_base64 = get_base64_image("2-c-Turkish-Technic-scaled.jpg")

# CSS ile Arka Planı Hangar Yapma ve Modern Yüksek Okunabilirlik Temalandırması
bg_css = ""
if hangar_base64:
    bg_css = f"""
    .stApp {{
        background-image: url("{hangar_base64}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
    }}
    [data-testid="stAppViewContainer"] {{
        background-color: rgba(15, 23, 42, 0.2) !important; /* Karartma çok hafifletildi, arka plan net */
        backdrop-filter: none !important; /* Blur tamamen kaldırıldı */
        -webkit-backdrop-filter: none !important;
    }}
    """

st.markdown(f"""
    <style>
    {bg_css}
    
    [data-testid="stHeader"] {{
        background: transparent !important;
    }}
    
    /* 1. ANA ÇALIŞMA ALANINI BEYAZ VE ŞIK BİR KART YAPMA */
    .block-container {{
        background-color: rgba(255, 255, 255, 0.96) !important;
        padding: 3rem !important;
        border-radius: 16px;
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.4);
        max-width: 1100px !important;
        margin-top: 4vh !important;
        margin-bottom: 4vh !important;
    }}
    
    /* 2. ANA ALANDAKİ TÜM YAZILARI KESKİN SİYAH/KOYU GRİ YAPMA */
    h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown, [data-testid="stWidgetLabel"] p {{
        color: #1e293b !important;
        text-shadow: none !important;
    }}
    
    /* 3. SOLDALİ TAB (SIDEBAR) BEYAZ ARKA PLAN VE SİYAH YAZILAR */
    [data-testid="stSidebar"] {{
        background-color: rgba(255, 255, 255, 0.98) !important;
        border-right: 1px solid rgba(0,0,0,0.1);
    }}
    [data-testid="stSidebar"] * {{
        color: #0f172a !important;
        text-shadow: none !important;
    }}
    
    /* Sidebar içindeki butonlar için soft gri zemin */
    [data-testid="stSidebar"] button {{
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
    }}
    [data-testid="stSidebar"] button:hover {{
        background-color: #e2e8f0 !important;
        border-color: #94a3b8 !important;
    }}
    
    /* 4. SEÇENEK SEÇİM EKRANINDAKİ (CHOOSE MODE) KUTUCUKLARIN BEYAZ VE BELİRGİN YAPILMASI */
    div[data-testid="column"] {{
        background-color: #ffffff !important;
        padding: 2.5rem !important;
        border-radius: 12px !important;
        border: 1px solid #e2e8f0 !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05) !important;
    }}
    
    /* Radyo butonu seçenek metinleri */
    .stRadio [data-testid="stMarkdownContainer"] {{
        font-size: 16px !important;
        font-weight: 500;
        color: #1e293b !important;
    }}
    
    /* Dosya yükleme dropzone'u */
    [data-testid="stUploadDropzone"] {{
        background-color: #f8fafc !important;
        border: 2px dashed #3b82f6 !important;
    }}
    [data-testid="stUploadDropzone"] span, [data-testid="stUploadDropzone"] p {{
        text-shadow: none !important;
        color: #475569 !important;
    }}
    
    /* Bilgilendirme ve uyarı kutuları (Alerts) */
    .stAlert {{
        background-color: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
    }}
    .stAlert p, .stAlert span, .stAlert div {{
        color: #1e293b !important;
        text-shadow: none !important;
    }}
    
    /* Progress Bar rengi (Mavi) */
    .stProgress > div > div > div > div {{
        background-color: #3b82f6 !important;
    }}
    
    /* Tablo veri hücresi gölgelerini kaldır */
    div[data-testid="stDataFrame"] * {{
        text-shadow: none !important;
    }}
    </style>
""", unsafe_allow_html=True)

# State (Durum) Yönetimi
if 'step' not in st.session_state:
    st.session_state.step = "upload"
if 'selected_jobs' not in st.session_state:
    st.session_state.selected_jobs = {}
if 'raw_data' not in st.session_state:
    st.session_state.raw_data = None
if 'unique_dates' not in st.session_state:
    st.session_state.unique_dates = []
if 'current_date_idx' not in st.session_state:
    st.session_state.current_date_idx = 0
if 'original_file_bytes' not in st.session_state:
    st.session_state.original_file_bytes = None
if 'original_filename' not in st.session_state:
    st.session_state.original_filename = "Logbook_Raporu.xlsx"

# Excel Verisini Yükleme Fonksiyonu (Standartlaştırılmış Tarih Formatı ile)
def load_excel_data(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    sheet = wb.active
    
    data = []
    # 1. ve 2. satırlar başlık, veri 3'ten başlar
    for r_idx in range(3, sheet.max_row + 1):
        row_vals = [sheet.cell(row=r_idx, column=c_idx).value for c_idx in range(1, sheet.max_column + 1)]
        if any(row_vals):
            date_val = row_vals[1]
            standard_date_str = ""
            if date_val:
                if hasattr(date_val, 'strftime'): 
                    standard_date_str = date_val.strftime('%d.%m.%Y')
                else: 
                    date_str = str(date_val).strip()
                    parsed_dt = None
                    for fmt in ('%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
                        try:
                            parsed_dt = pd.to_datetime(date_str, format=fmt)
                            break
                        except:
                            continue
                    if parsed_dt is not None and not pd.isna(parsed_dt):
                        standard_date_str = parsed_dt.strftime('%d.%m.%Y')
                    else:
                        standard_date_str = date_str # Fallback
            
            data.append({
                "row_idx": r_idx,
                "no": row_vals[0],
                "date": standard_date_str,
                "location": row_vals[2],
                "fleet": row_vals[3],
                "reg": row_vals[7],
                "description": row_vals[22],
                "duration": row_vals[23],
                "ref": row_vals[24],
                "wo": row_vals[25]
            })
    
    df = pd.DataFrame(data)
    df['date'] = df['date'].astype(str).str.strip()
    return df

# Formatı %100 Koruyan Excel Çıktısı Üretme Fonksiyonu
def generate_output_excel(original_bytes, selected_row_indices, yellow_row_indices):
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
    sheet = wb.active
    
    # 1. ADIM: Sarı boyanacak satırları boya
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for r_idx in yellow_row_indices:
        for col in range(1, sheet.max_column + 1):
            sheet.cell(row=r_idx, column=col).fill = yellow_fill

    # 2. ADIM: Seçilmeyen işleri tablodan sil (Aşağıdan yukarıya doğru)
    max_row = sheet.max_row
    for r_idx in range(max_row, 2, -1):
        if r_idx not in selected_row_indices:
            sheet.delete_rows(r_idx)
            
    # 3. ADIM: Sıra No kolonunu (1. Sütun) yeniden düzenle
    current_no = 1
    for r_idx in range(3, sheet.max_row + 1):
        sheet.cell(row=r_idx, column=1, value=current_no)
        current_no += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ----------------- ADIM 1: DOSYA YÜKLEME -----------------
if st.session_state.step == "upload":
    st.title("✈️ Logbook Düzenleme Otomasyonu")
    st.subheader("Orijinal Excel (.xlsx) Dosyanızı Yükleyin")
    
    if not hangar_base64:
        st.warning("İpucu: Arka planın hangar resmi olması için '2-c-Turkish-Technic-scaled.jpg' dosyasını bu proje klasörüne kopyalayın.")
    if not plane_base64:
        st.warning("İpucu: Sonda uçağın uçması için 'pngegg.jpg' dosyasını bu proje klasörüne kopyalayın.")
        
    uploaded_file = st.file_uploader("Dosya Seçin", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.getvalue()
            df = load_excel_data(file_bytes)
            
            # Tarihleri kronolojik (Yıl -> Ay -> Gün) olarak sıralama mekanizması
            parsed_dates = []
            for d in df['date'].unique():
                try:
                    dt = pd.to_datetime(d, format='%d.%m.%Y')
                    parsed_dates.append((dt, d))
                except:
                    parsed_dates.append((pd.Timestamp.max, d))
            
            parsed_dates.sort(key=lambda x: x[0])
            sorted_unique_dates = [x[1] for x in parsed_dates]
            
            st.session_state.original_file_bytes = file_bytes
            st.session_state.original_filename = uploaded_file.name
            st.session_state.raw_data = df
            st.session_state.unique_dates = sorted_unique_dates
            st.session_state.selected_jobs = {}
            st.session_state.step = "choose_mode"
            st.rerun()
        except Exception as e:
            st.error(f"Hata: {e}. Lütfen doğru şablonda bir .xlsx dosyası yükleyin.")

# ----------------- ADIM 2: SEÇİM MODU -----------------
elif st.session_state.step == "choose_mode":
    st.title("⚙️ İşlem Modunu Seçin")
    st.write("Dosyanız başarıyla yüklendi. İşlemleri nasıl yapmak istersiniz?")
    
    col1, col2 = st.columns(2)
    with col1:
        st.success("⚡ 1. Seçenek: Otomatik (Hızlı)")
        st.write("**Otomatik ilk satırı seçer diğerlerini siler.**")
        st.write("Aynı günde sadece 1 iş seçimi yapar. Sisteme yüklenen listedeki her gün için ilk sıradaki işi otomatik tutar, geri kalanları siler ve sizi doğrudan Sarı Boyama ekranına yönlendirir.")
        if st.button("Otomatik Seçimle İlerle", use_container_width=True, type="primary"):
            for date in st.session_state.unique_dates:
                daily_jobs = st.session_state.raw_data[st.session_state.raw_data['date'] == date]
                first_row_idx = daily_jobs.iloc[0]['row_idx']
                st.session_state.selected_jobs[date] = first_row_idx
            
            st.session_state.step = "select_samples"
            st.rerun()
            
    with col2:
        st.info("🖐️ 2. Seçenek: Manuel (Kontrollü)")
        st.write("**Tek tek elle seçerek devam edilir.**")
        st.write("Tüm günler için o güne ait işler listelenir. Hangi işin tabloda kalacağına okuyarak bizzat siz karar verirsiniz.")
        if st.button("Manuel Seçime Başla", use_container_width=True):
            st.session_state.current_date_idx = 0
            st.session_state.step = "select_daily"
            st.rerun()

# ----------------- ADIM 3: MANUEL GÜNLÜK İŞ SEÇİMİ (KLAVYE DESTEKLİ) -----------------
elif st.session_state.step == "select_daily":
    df = st.session_state.raw_data
    dates = st.session_state.unique_dates
    current_idx = st.session_state.current_date_idx
    current_date = dates[current_idx]
    
    # --- JAVASCRIPT: OK TUŞLARI + ENTER NAVİGASYONU VE SIDEBAR SCROLL ---
    components.html("""
    <script>
    const doc = window.parent.document;
    
    // 1. Sol menüdeki aktif günü ortala
    const sidebar = doc.querySelector('[data-testid="stSidebar"]');
    if (sidebar) {
        const allElements = Array.from(sidebar.querySelectorAll('*'));
        const activeElement = allElements.find(el => el.textContent && el.textContent.includes('👉'));
        if (activeElement) {
            activeElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }

    // 2. Klavye Dinleyicisi
    if (!window.parent._keyboardListenerAdded) {
        window.parent._keyboardListenerAdded = true;
        doc.addEventListener('keydown', function(e) {
            if (e.target.tagName === 'INPUT' && e.target.type === 'text') return;
            
            const radioGroup = doc.querySelector('div[data-testid="stRadio"]');
            
            if (radioGroup) {
                const radios = Array.from(radioGroup.querySelectorAll('input[type="radio"]'));
                const checkedIndex = radios.findIndex(r => r.checked);
                
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    if (checkedIndex < radios.length - 1) {
                        radios[checkedIndex + 1].click();
                    }
                } 
                else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (checkedIndex > 0) {
                        radios[checkedIndex - 1].click();
                    }
                }
            }
            
            if (e.key === 'Enter') {
                const buttons = Array.from(doc.querySelectorAll('button'));
                const nextBtn = buttons.find(btn => btn.textContent.includes('Sonraki Gün') || btn.textContent.includes('Örnek İş Seçimine Geç'));
                if (nextBtn) {
                    nextBtn.click();
                }
            }
        });
    }
    </script>
    """, height=0)

    # --- KRONOLOJİK GRUPLANDIRMA (SIDEBAR HAZIRLIĞI) ---
    grouped = defaultdict(lambda: defaultdict(list))
    for idx, d_str in enumerate(dates):
        try:
            parts = d_str.split('.')
            day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
        except:
            year, month = 9999, 12
        grouped[year][month].append((idx, d_str))
        
    MONTH_NAMES = {
        1: 'Ocak', 2: 'Şubat', 3: 'Mart', 4: 'Nisan', 5: 'Mayıs', 6: 'Haziran',
        7: 'Temmuz', 8: 'Ağustos', 9: 'Eylül', 10: 'Ekim', 11: 'Kasım', 12: 'Aralık'
    }

    # Aktif günü bulup yılına ve ayına göre expander durumunu ayarla
    try:
        parts = current_date.split('.')
        active_year = int(parts[2])
        active_month = int(parts[1])
    except:
        active_year = None
        active_month = None

    # --- SOL SÜTUN (YANDAKİ NAVİGASYON - NESTED EXPANDERS) ---
    st.sidebar.title("📅 Gün Seçici")
    
    for year in sorted(grouped.keys()):
        is_year_expanded = (year == active_year)
        # Yıl Expander'ı (Sidebar içinde)
        with st.sidebar.expander(f"📁 {year} Yılı", expanded=is_year_expanded):
            for month in sorted(grouped[year].keys()):
                is_month_expanded = (year == active_year and month == active_month)
                # Ay Expander'ı (Yıl expander'ının İÇİNDE - NOT: st.sidebar kullanılmaz)
                with st.expander(f"📅 {MONTH_NAMES.get(month, month)}", expanded=is_month_expanded):
                    for idx, d_str in grouped[year][month]:
                        status_icon = "🟢" if d_str in st.session_state.selected_jobs else "⚪"
                        
                        # Butonları çizerken st.sidebar.button DEĞİL, normal st.button kullanıyoruz.
                        # Böylece elemanlar sidebar'ın en altına uçmak yerine tam olarak Ay klasörünün içine yerleşir.
                        if idx == current_idx:
                            label = f"👉 {d_str} (Aktif)"
                            if st.button(label, key=f"nav_btn_{d_str}", use_container_width=True, type="primary"):
                                st.session_state.current_date_idx = idx
                                st.rerun()
                        else:
                            label = f"{status_icon} {d_str}"
                            if st.button(label, key=f"nav_btn_{d_str}", use_container_width=True):
                                st.session_state.current_date_idx = idx
                                st.rerun()

    # --- SAĞ SÜTUN (ANA EKRAN) ---
    st.title("⚡ Seri İş Seçim Ekranı")
    progress = len(st.session_state.selected_jobs) / len(dates)
    st.progress(progress)
    st.subheader(f"Tarih: {current_date} ({current_idx + 1} / {len(dates)})")
    
    st.info("⌨️ **Klavye Kısayolları:** Seçenekler arasında gezmek için **Aşağı / Yukarı Ok** tuşlarını kullanın. Seçimi onaylayıp sonraki güne geçmek için **Enter** tuşuna basın.")
    
    daily_jobs = df[df['date'] == current_date]
    
    options = {}
    for _, row in daily_jobs.iterrows():
        label = f"W/O: {row['wo']} | Ref: {row['ref']} | Süre: {row['duration']} | Tanım: {row['description']}"
        options[row['row_idx']] = label
        
    default_val = list(options.keys())[0]
    if current_date in st.session_state.selected_jobs:
        default_val = st.session_state.selected_jobs[current_date]
        
    selected_row = st.radio(
        "Kayıt listesinden kalacak olan 1 işi seçin:",
        options=list(options.keys()),
        format_func=lambda x: options[x],
        index=list(options.keys()).index(default_val),
        key=f"radio_{current_date}"
    )
    
    st.session_state.selected_jobs[current_date] = selected_row
    st.write("---")
    
    col_nav1, col_nav2 = st.columns([1, 1])
    with col_nav1:
        if current_idx > 0:
            if st.button("⬅️ Önceki Gün", use_container_width=True):
                st.session_state.current_date_idx -= 1
                st.rerun()
                
    with col_nav2:
        if current_idx == len(dates) - 1 and len(st.session_state.selected_jobs) == len(dates):
            if st.button("🎉 Örnek İş Seçimine Geç (Enter)", type="primary", use_container_width=True):
                st.session_state.step = "select_samples"
                st.rerun()
        else:
            if st.button("Sonraki Gün ➡️ (Enter)", type="primary", use_container_width=True):
                st.session_state.current_date_idx += 1
                st.rerun()

# ----------------- ADIM 4: ÖRNEK İŞ SEÇİMİ (SARI BOYAMA) -----------------
elif st.session_state.step == "select_samples":
    st.title("🎯 Örnek İşlerin Belirlenmesi")
    
    selected_indices = list(st.session_state.selected_jobs.values())
    df = st.session_state.raw_data
    selected_df = df[df['row_idx'].isin(selected_indices)].copy()
    
    selected_df['temp_sort_date'] = pd.to_datetime(selected_df['date'], format='%d.%m.%Y', errors='coerce')
    selected_df = selected_df.sort_values('temp_sort_date')
    
    st.write(f"Toplam **{len(selected_df)}** gün/iş kronolojik olarak filtrelendi.")
    st.info("Otoriteye örnek olarak gösterilecek (sarıya boyanacak) işleri sol taraftaki kutucuklardan seçin.")
    
    selected_df['Sarı Boya (Örnek İş)'] = False
    
    edited_df = st.data_editor(
        selected_df[['Sarı Boya (Örnek İş)', 'date', 'location', 'fleet', 'reg', 'description', 'duration', 'wo', 'row_idx']],
        disabled=['date', 'location', 'fleet', 'reg', 'description', 'duration', 'wo', 'row_idx'],
        column_config={
            "Sarı Boya (Örnek İş)": st.column_config.CheckboxColumn("Sarı Boya?", default=False)
        },
        hide_index=True,
        use_container_width=True
    )
    
    yellow_rows = edited_df[edited_df['Sarı Boya (Örnek İş)'] == True]['row_idx'].tolist()
    st.metric("Seçilen Örnek İş Sayısı (Sarı Boyanacaklar)", f"{len(yellow_rows)}")
    
    col_b1, col_b2 = st.columns([1, 1])
    with col_b1:
        if st.button("⬅️ Geri Dön ve Mod Seç", use_container_width=True):
            st.session_state.step = "choose_mode"
            st.rerun()
            
    with col_b2:
        if st.button("Excel Dosyasını Üret ve Hazırla 📥", type="primary", use_container_width=True):
            with st.spinner("Şablon korunarak Excel oluşturuluyor..."):
                excel_data = generate_output_excel(
                    st.session_state.original_file_bytes, 
                    selected_indices, 
                    yellow_rows
                )
                st.session_state.final_excel = excel_data
                st.session_state.step = "download"
                st.rerun()

# ----------------- ADIM 5: İNDİRME VE SAĞDAN SOLA THY UÇAĞI ANİMASYONU EKRANI -----------------
elif st.session_state.step == "download":
    st.title("🏆 Raporunuz Hazır!")
    
    # --- JAVASCRIPT: GERÇEK UÇAK RESMİ İLE SAĞDAN SOLA UÇUŞ ANİMASYONU ---
    if plane_base64:
        plane_html = f'<img class="thy-plane-img" src="{plane_base64}">'
        plane_css = """
        .thy-plane-img {
            width: 380px;
            height: auto;
            filter: drop-shadow(-10px 15px 15px rgba(0,0,0,0.5));
        }
        """
    else:
        plane_html = '<span style="font-size: 130px; filter: drop-shadow(5px 10px 10px rgba(0,0,0,0.3)); display: inline-block; transform: scaleX(-1);">✈️</span><span style="font-size: 50px; font-weight: 900; font-style: italic; color: #D61C22; background: white; padding: 5px 20px; border-radius: 12px; border: 4px solid #D61C22; margin-left: 15px; box-shadow: -3px 5px 15px rgba(0,0,0,0.4);">THY</span>'
        plane_css = ""

    components.html(f"""
    <script>
    const parentDoc = window.parent.document;
    if (!parentDoc.getElementById('thy-anim')) {{
        const style = parentDoc.createElement('style');
        style.innerHTML = `
        @keyframes flyPlane {{
            0% {{ left: 120vw; top: 75vh; transform: rotate(15deg) scale(0.6); opacity: 0; }}
            10% {{ opacity: 1; }}
            90% {{ opacity: 1; }}
            100% {{ left: -40vw; top: 10vh; transform: rotate(5deg) scale(1.3); opacity: 0; }}
        }}
        .thy-plane-wrapper {{
            position: fixed;
            z-index: 9999999;
            pointer-events: none;
            animation: flyPlane 5.5s cubic-bezier(0.25, 1, 0.5, 1) forwards;
            display: flex;
            align-items: center;
        }}
        {plane_css}
        `;
        parentDoc.head.appendChild(style);
        
        const plane = parentDoc.createElement('div');
        plane.id = 'thy-anim';
        plane.className = 'thy-plane-wrapper';
        plane.innerHTML = '{plane_html}';
        parentDoc.body.appendChild(plane);
        
        setTimeout(() => {{
            if (plane.parentNode) plane.parentNode.removeChild(plane);
        }}, 6000);
    }}
    </script>
    """, height=0)

    st.success("Tebrikler! Seçtiğiniz ayarlara göre Excel dosyası, şablon yapısı %100 korunarak başarıyla oluşturuldu.")
    
    st.download_button(
        label="📥 Düzenlenmiş Excel Dosyasını İndir (.xlsx)",
        data=st.session_state.final_excel,
        file_name=st.session_state.original_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    if st.button("🔄 Yeni Bir Dosya İle Yeniden Başla", use_container_width=True):
        st.session_state.clear()
        st.rerun()
