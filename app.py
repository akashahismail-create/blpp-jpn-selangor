# BLPP JPN Selangor - v50 PRO UPGRADE - 504 Pusat + Amali Sains + Logo LP Official
# Upgrade dari v41 FINAL by Akashah Ismail - 2026
import streamlit as st
import pandas as pd
import json, os, base64
from datetime import datetime, date
from io import BytesIO
from PIL import Image
import io

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except:
    GSHEETS_AVAILABLE = False

# ========== CONFIG ==========
st.set_page_config(
    page_title="BLPP JPN Selangor v50 PRO - LP Official",
    layout="wide",
    page_icon="📋",
    initial_sidebar_state="expanded"
)

# ========== LOGO HANDLING ==========
def get_logo_base64():
    for p in ["/mnt/data/logo_lp_transparent.png", "/mnt/data/logo_lp.png", "logo_lp.png"]:
        if os.path.exists(p):
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode()
    return None

LOGO_B64 = get_logo_base64()

# ========== MODERN CSS v50 ==========
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=Plus+Jakarta+Sans:wght@700;800&display=swap');
* {{ font-family: 'Inter', sans-serif; }}
#MainMenu, header, footer, [data-testid="stToolbar"], [data-testid="stHeader"] {{visibility: hidden; display: none;}}
.stApp {{ background: linear-gradient(180deg, #F8FAFF 0%, #EEF2FF 100%) !important; }}

/* FORCE LIGHT MODE - FIX HP DARK MODE */
html, body, .stApp {{ color-scheme: light !important; }}
div[data-testid="stSelectbox"] div[data-baseweb="select"],
div[data-testid="stSelectbox"] div[data-baseweb="select"] *,
div[data-testid="stTextInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-testid="stDateInput"] input,
input, textarea, select {{
    color: #000000 !important;
    background-color: #FFFFFF !important;
    -webkit-text-fill-color: #000000 !important;
}}
div[data-baseweb="popover"], div[data-baseweb="menu"], ul[data-baseweb="menu"], li[data-baseweb="option"] {{
    background-color: #FFFFFF !important;
    color: #000000 !important;
}}
li[data-baseweb="option"] * {{ color: #000000 !important; }}

/* HEADER v50 PRO - DENGAN LOGO LP */
.header-pro {{
    background: linear-gradient(135deg, #07112A 0%, #0A1931 20%, #0D47A1 60%, #1565C0 100%);
    padding:26px 28px; border-radius:20px; margin-bottom:20px; text-align:center;
    box-shadow:0 14px 40px rgba(13,71,161,0.3); border:1px solid rgba(255,255,255,0.12);
    position: relative; overflow: hidden;
}}
.header-pro::before {{
    content:''; position:absolute; top:-50%; left:-20%; width:140%; height:200%;
    background: radial-gradient(ellipse at center, rgba(255,255,255,0.08) 0%, transparent 60%);
}}
.header-pro img {{
    height:58px; background:white; padding:6px 18px; border-radius:12px; margin-bottom:14px;
    box-shadow:0 4px 16px rgba(0,0,0,0.2);
}}
.header-pro h1 {{ color:white; margin:0; font-family:'Plus Jakarta Sans'; font-size:24px; font-weight:800; letter-spacing:-0.5px; position:relative; }}
.header-pro h2 {{ color:#FFECB3; margin:8px 0 0 0; font-size:12px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; position:relative; }}
.header-pro h3 {{ color:white; margin:14px 0 0 0; font-size:14px; font-weight:700; background:rgba(255,255,255,0.18); display:inline-block; padding:8px 22px; border-radius:24px; backdrop-filter:blur(10px); position:relative; }}
.header-pro .ver {{ position:absolute; top:14px; right:18px; background:#FFD54F; color:#0A1931; font-size:10px; font-weight:800; padding:4px 10px; border-radius:20px; }}

.card {{ background:white; border-radius:16px; padding:20px; border:1px solid #E8EEFF; box-shadow:0 6px 20px rgba(0,0,0,0.06); margin:14px 0; transition:0.2s; }}
.card:hover {{ box-shadow:0 10px 28px rgba(0,0,0,0.09); transform:translateY(-1px); }}
.card h3 {{ font-family:'Plus Jakarta Sans'; font-weight:800; font-size:16px; margin:0 0 8px 0; color:#0A1931; }}
.metric-pro {{ background:white; border-radius:14px; padding:16px; border-left:4px solid #1565C0; box-shadow:0 4px 12px rgba(0,0,0,0.05); }}

div[data-testid="stButton"] button {{ border-radius:12px !important; font-weight:700 !important; padding:8px 18px !important; }}
</style>
""", unsafe_allow_html=True)

# ========== DATA PPD ==========
PPD_LIST = ["Petaling Utama","Petaling Perdana","Gombak","Hulu Langat","Klang","Kuala Langat","Kuala Selangor","Hulu Selangor","Sepang","Sabak Bernam"]
KOD_PPD_MAPPING = {"BA":"Klang","BB":"Kuala Langat","BC":"Kuala Selangor","BD":"Hulu Langat","BE":"Hulu Selangor","BF":"Sabak Bernam","BG":"Gombak","BH":"Petaling Perdana","BJ":"Sepang","BK":"Petaling Utama"}
PPD_TO_KODS_STRICT = {"Klang":["BA"],"Kuala Langat":["BB"],"Kuala Selangor":["BC"],"Hulu Langat":["BD"],"Hulu Selangor":["BE"],"Sabak Bernam":["BF"],"Gombak":["BG"],"Petaling Perdana":["BH"],"Sepang":["BJ"],"Petaling Utama":["BK"]}

def get_kods_for_ppd(ppd_name):
    if not ppd_name: return []
    if ppd_name in PPD_TO_KODS_STRICT: return PPD_TO_KODS_STRICT[ppd_name]
    for k,v in PPD_TO_KODS_STRICT.items():
        if k.lower() in str(ppd_name).lower(): return v
    return []

def is_report_for_ppd_strict(report, target_ppd):
    kods = get_kods_for_ppd(target_ppd)
    no_pusat = str(report.get("no_pusat","")).upper()
    return any(no_pusat.startswith(k) for k in kods)

# ========== LOAD DATA EXCEL - NEW v50 ==========
@st.cache_data
def load_spm_data():
    path = "/mnt/data/data_v1.1_3_1.xlsx"
    if not os.path.exists(path): return None, None, None
    try:
        pusat = pd.read_excel(path, sheet_name="selenggara_pusat")
        mp = pd.read_excel(path, sheet_name="Data_Lengkap_4Pakej")
        bilik = pd.read_excel(path, sheet_name="Senarai_Bilik_Kebal")
        return pusat, mp, bilik
    except Exception as e:
        st.error(f"Gagal load SPM data: {e}")
        return None,None,None

@st.cache_data
def load_amali_data():
    path = "/mnt/data/data_pusat_amali_sains_baru_1.xlsx"
    if not os.path.exists(path): return None, None
    try:
        ringkas = pd.read_excel(path, sheet_name="Ringkasan_Makmal")
        detail = pd.read_excel(path, sheet_name="Data_Pusat_Amali_Sains")
        return ringkas, detail
    except Exception as e:
        st.error(f"Gagal load Amali data: {e}")
        return None,None

# ========== SIMULASI DB ==========
def load_users():
    # Cuba cari file users.json atau guna default demo
    if os.path.exists("users.json"):
        with open("users.json","r") as f: return json.load(f)
    # Generate dari Excel pusat
    pusat,_,_ = load_spm_data()
    users=[]
    if pusat is not None:
        for idx,row in pusat.iterrows():
            no = str(row.get("No_Pusat",""))
            kod = no[:2] if len(no)>=2 else "BA"
            ppd = KOD_PPD_MAPPING.get(kod,"Klang")
            users.append({
                "id": no, "nama": f"KP {no}", "peranan":"KP","ppd":ppd,
                "no_pusat":no,"nama_pusat":str(row.get("Nama_Pusat","")),"password":no.lower()
            })
    # Tambah PPD/JPN/LP default
    users.extend([
        {"id":"PPD01","nama":"PPD Petaling Utama","peranan":"PPD","ppd":"Petaling Utama","no_pusat":"","nama_pusat":"","password":"ppd01"},
        {"id":"JPN001","nama":"JPN Selangor","peranan":"ADMIN_JPN","ppd":"","no_pusat":"","nama_pusat":"","password":"jpn123"},
        {"id":"LP001","nama":"Lembaga Peperiksaan","peranan":"LP","ppd":"","no_pusat":"","nama_pusat":"","password":"lp123"},
    ])
    return users

def save_users(u):
    with open("users.json","w") as f: json.dump(u,f,indent=2)

def load_reports():
    if os.path.exists("reports.json"):
        with open("reports.json","r") as f: return json.load(f)
    return []

def save_reports(r):
    with open("reports.json","w") as f: json.dump(r,f,indent=2,ensure_ascii=False)

# ========== HEADER COMPONENT v50 ==========
def render_header(user=None):
    logo_html = f'<img src="data:image/png;base64,{LOGO_B64}" alt="LP Logo">' if LOGO_B64 else '<div style="font-size:40px;">📋</div>'
    role_text = f"{user['peranan']} - {user.get('ppd','')} {user.get('no_pusat','')}" if user else "SISTEM BLPP RASMI LP/Am 90 Pin. 1/2023"
    st.markdown(f"""
    <div class="header-pro">
        <div class="ver">v50 PRO UPGRADE</div>
        {logo_html}
        <h1>Sistem Pengurusan Pusat Peperiksaan</h1>
        <h2>Borang Laporan Penyelewengan / Bencana (BLPP) • Lembaga Peperiksaan Malaysia</h2>
        <h3>{role_text}</h3>
    </div>
    """, unsafe_allow_html=True)

# ========== PDF GENERATOR WITH LOGO ==========
def jana_blpp_html(report, logo_b64):
    # Simplified official template with logo
    kes_rows = ""
    if report.get("tiada_kes"):
        kes_rows = "<tr><td colspan='5' style='text-align:center;padding:20px;'>TIADA KES / TIADA PENYELEWENGAN</td></tr>"
    else:
        for i,k in enumerate(report.get("kes_data",[]),1):
            kes_rows += f"<tr><td>{i}</td><td>{k.get('jenis','')}</td><td>{k.get('kertas','')}</td><td>{k.get('bil_calon','')}</td><td>{k.get('catatan','')}</td></tr>"
    
    logo_img = f'<img src="data:image/png;base64,{logo_b64}" style="height:55px;">' if logo_b64 else ''
    html = f"""
    <html><head><style>
    body {{ font-family: Arial; font-size:11px; color:#000; }}
    .kop {{ text-align:center; border-bottom:3px double #000; padding-bottom:10px; margin-bottom:12px; }}
    table {{ width:100%; border-collapse:collapse; }}
    th, td {{ border:1px solid #000; padding:6px; font-size:10px; }}
    th {{ background:#E8EEFF; }}
    </style></head><body>
    <div class="kop">{logo_img}<br><b>LEMBAGA PEPERIKSAAN MALAYSIA</b><br>BORANG LAPORAN PENYELEWENGAN / BENCANA (BLPP)<br>LP/Am 90 Pin. 1/2023</div>
    <table>
    <tr><th>Tarikh</th><td>{report.get('tarikh','')}</td><th>No Pusat</th><td>{report.get('no_pusat','')}</td></tr>
    <tr><th>Nama Pusat</th><td colspan="3">{report.get('nama_pusat','')}</td></tr>
    <tr><th>Peperiksaan</th><td>{report.get('peperiksaan','')}</td><th>PPD</th><td>{report.get('ppd','')}</td></tr>
    </table><br>
    <table><tr><th>Bil</th><th>Jenis Kes</th><th>Kertas</th><th>Bil Calon</th><th>Catatan</th></tr>{kes_rows}</table>
    <br><br><table style="border:none;"><tr style="border:none;"><td style="border:none;width:60%"></td><td style="border:none;text-align:center;">Ketua Pengawas<br><br><br>({report.get('nama_kp','')})<br>{report.get('no_pusat','')}</td></tr></table>
    </body></html>
    """
    return html

# ========== LOGIN ==========
if "user" not in st.session_state:
    st.session_state.user = None

if not st.session_state.user:
    render_header()
    col1,col2,col3 = st.columns([1,1.2,1])
    with col2:
        st.markdown('<div class="card"><h3>🔐 Log Masuk Sistem</h3><p style="font-size:12px;color:#64748B;">Guna ID Pusat sebagai login: BH562 / bh562</p></div>', unsafe_allow_html=True)
        with st.form("login"):
            id_input = st.text_input("ID Pengguna", placeholder="Contoh: BH562 / PPD01 / JPN001")
            pass_input = st.text_input("Kata Laluan", type="password")
            peranan = st.selectbox("Peranan", ["KP","PPD","ADMIN_JPN","LP"])
            login_btn = st.form_submit_button("Masuk Sistem →", use_container_width=True)
            if login_btn:
                users = load_users()
                found = None
                for u in users:
                    if u["id"].lower()==id_input.lower() and u["password"]==pass_input and u["peranan"]==peranan:
                        found=u; break
                    # flexible: id as no_pusat also
                    if u.get("no_pusat","").lower()==id_input.lower() and u["password"]==pass_input:
                        found=u; break
                if found:
                    st.session_state.user = found
                    st.rerun()
                else:
                    st.error("ID / Password / Peranan salah. Cuba BH562 / bh562 sebagai KP")
        st.markdown('<div style="text-align:center;font-size:11px;color:#94A3B8;margin-top:20px;">Sistem Rasmi JPN Selangor • 504 Pusat • Data SPM & Amali Sains Integrated • v50 PRO</div>', unsafe_allow_html=True)
    st.stop()

# ========== DASHBOARD UTAMA ==========
user = st.session_state.user
render_header(user)

# Sidebar menu - still in main
col_menu, col_content = st.columns([1,3.2])

with col_menu:
    st.markdown(f'<div class="card" style="text-align:center;"><b>{user["nama"]}</b><br><span style="font-size:12px;color:#64748B;">{user["id"]} • {user["peranan"]}</span><br><span style="font-size:11px;background:#EEF2FF;padding:4px 10px;border-radius:20px;">{user.get("ppd","")}</span></div>', unsafe_allow_html=True)
    menu = st.radio("Menu Utama", ["📝 Isi Borang BLPP", "🏫 Data Pusat & Amali [BARU]", "📋 Senarai Laporan", "🖨️ Cetak Borang", "📊 Statistik JPN", "⚙️ Pengurusan"], label_visibility="collapsed")
    if st.button("Keluar Log", use_container_width=True):
        st.session_state.user=None
        st.rerun()

reports = load_reports()

with col_content:
    if menu == "📝 Isi Borang BLPP":
        st.markdown('<div class="card"><h3>📝 Borang BLPP - LP/Am 90 Pin. 1/2023</h3><p style="font-size:12px;">Sila isi dengan lengkap. Logo LP rasmi akan tercetak pada borang.</p></div>', unsafe_allow_html=True)
        # Quick prefill
        with st.form("blpp_form"):
            c1,c2,c3 = st.columns(3)
            with c1:
                tarikh = st.date_input("Tarikh Peperiksaan", value=date.today())
                no_pusat = st.text_input("No Pusat", value=user.get("no_pusat",""))
            with c2:
                nama_pusat = st.text_input("Nama Pusat", value=user.get("nama_pusat",""))
                peperiksaan = st.selectbox("Peperiksaan", ["SPM","SPM Ulangan","STAM","STPM"])
            with c3:
                tahun = st.text_input("Tahun", value="2026")
                ppd = st.text_input("PPD", value=user.get("ppd",""))

            st.markdown("---")
            tiada_kes = st.checkbox("☑ TIADA KES / TIADA PENYELEWENGAN BERLAKU")
            
            kes_data=[]
            if not tiada_kes:
                st.markdown("**Butiran Kes (boleh tambah 3 kes)**")
                for i in range(3):
                    cols = st.columns([1.5,1,0.8,2])
                    with cols[0]: jenis = st.selectbox(f"Jenis Kes {i+1}", ["","Kertas Bocor","Salah Edar Kertas","Calon Meniru","Kecemasan Bencana","Lain-lain"], key=f"j{i}")
                    with cols[1]: kertas = st.text_input(f"Kertas {i+1}", key=f"k{i}")
                    with cols[2]: bil = st.number_input(f"Bil Calon {i+1}", min_value=0, key=f"b{i}")
                    with cols[3]: catatan = st.text_input(f"Catatan {i+1}", key=f"c{i}")
                    if jenis: kes_data.append({"jenis":jenis,"kertas":kertas,"bil_calon":bil,"catatan":catatan})
            
            nama_kp = st.text_input("Nama Ketua Pengawas", value=user.get("nama",""))
            telefon = st.text_input("No Telefon KP")
            hantar = st.form_submit_button("✅ Hantar Laporan Rasmi + Jana PDF", use_container_width=True)
            
            if hantar:
                new_report = {
                    "id": f"BLPP-{no_pusat}-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "tarikh": str(tarikh),
                    "tarikh_submit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "peperiksaan": peperiksaan, "tahun": tahun,
                    "no_pusat": no_pusat, "nama_pusat": nama_pusat,
                    "nama_kp": nama_kp, "ppd": ppd, "id_kp": user["id"],
                    "kes_data": kes_data, "tiada_kes": tiada_kes,
                    "telefon_kp": telefon
                }
                reports.append(new_report)
                save_reports(reports)
                st.success(f"✅ Laporan {new_report['id']} berjaya dihantar!")
                html = jana_blpp_html(new_report, LOGO_B64)
                st.components.v1.html(html, height=600, scrolling=True)
                st.download_button("📥 Download Borang HTML (Cetak PDF)", html, file_name=f"{new_report['id']}.html", mime="text/html")

    elif menu == "🏫 Data Pusat & Amali [BARU]":
        st.markdown('<div class="card"><h3>🏫 Modul Baru v50 - Data Pusat SPM & Amali Sains Integration</h3><p style="font-size:12px;">Gabungan data_v1.1 (4 Pakej, Bilik Kebal) + data Amali Sains (Makmal, Sidang)</p></div>', unsafe_allow_html=True)
        pusat_df, mp_df, bilik_df = load_spm_data()
        ringkas_df, detail_df = load_amali_data()

        tab1, tab2, tab3, tab4 = st.tabs(["🏫 Senarai Pusat (504)", "📦 Pakej Kertas (4 Pakej)", "🧪 Makmal Amali Sains", "🔍 Carian Pusat"])
        
        with tab1:
            if pusat_df is not None:
                # Filter ikut PPD user
                if user["peranan"]=="PPD":
                    kods = get_kods_for_ppd(user["ppd"])
                    filtered = pusat_df[pusat_df["No_Pusat"].astype(str).str[:2].isin(kods)]
                elif user["peranan"]=="KP":
                    filtered = pusat_df[pusat_df["No_Pusat"]==user.get("no_pusat","")]
                else:
                    filtered = pusat_df
                st.metric("Jumlah Pusat", len(filtered))
                st.dataframe(filtered, use_container_width=True, height=500)
                out = BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as w: filtered.to_excel(w,index=False)
                st.download_button("📥 Download Excel Pusat", out.getvalue(), file_name="Pusat_SPM.xlsx")
            else:
                st.warning("Data pusat belum ada")
        
        with tab2:
            if mp_df is not None:
                st.markdown("**Data Lengkap 4 Pakej - Bil Calon, Bil Naskah, Bilik Kebal**")
                if user["peranan"]=="KP":
                    mp_filtered = mp_df[mp_df["No_Pusat"]==user.get("no_pusat","")]
                else:
                    mp_filtered = mp_df.head(200)
                st.dataframe(mp_filtered, use_container_width=True, height=500)
                # Summary by Pakej
                st.bar_chart(mp_filtered.groupby("Pakej").size() if not mp_filtered.empty else pd.DataFrame())
            else:
                st.info("Tiada data pakej")
        
        with tab3:
            if ringkas_df is not None:
                st.markdown("**Ringkasan Makmal - Fizik, Kimia, Biologi, Sains Tambahan**")
                if user["peranan"]=="PPD":
                    kods = get_kods_for_ppd(user["ppd"])
                    r_f = ringkas_df[ringkas_df["No_Pusat"].astype(str).str[:2].isin(kods)]
                elif user["peranan"]=="KP":
                    r_f = ringkas_df[ringkas_df["No_Pusat"]==user.get("no_pusat","")]
                else:
                    r_f = ringkas_df
                st.dataframe(r_f, use_container_width=True, height=500)
                colA,colB,colC = st.columns(3)
                with colA: st.metric("Total Makmal", len(r_f))
                with colB: st.metric("Total Calon Amali", int(r_f["Jumlah_Calon"].sum()) if not r_f.empty else 0)
                with colC: st.metric("Avg Kapasiti", f"{r_f['Kapasiti_Makmal'].mean():.0f}" if not r_f.empty else "0")
                
                if detail_df is not None:
                    st.markdown("--- **Detail Sidang per Subjek**")
                    st.dataframe(detail_df.head(300), use_container_width=True)
            else:
                st.warning("Data Amali Sains belum ada")
        
        with tab4:
            q = st.text_input("🔍 Cari No Pusat / Nama Sekolah", placeholder="Contoh: BA001 / SMK...")
            if q and pusat_df is not None:
                res = pusat_df[pusat_df.astype(str).apply(lambda x: x.str.contains(q, case=False, na=False)).any(axis=1)]
                st.dataframe(res, use_container_width=True)
                # Show related amali + pakej
                if not res.empty and ringkas_df is not None:
                    no = res.iloc[0]["No_Pusat"]
                    st.markdown(f"**Makmal untuk {no}**")
                    st.dataframe(ringkas_df[ringkas_df["No_Pusat"]==no], use_container_width=True)
                    if mp_df is not None:
                        st.markdown(f"**Pakej Kertas untuk {no}**")
                        st.dataframe(mp_df[mp_df["No_Pusat"]==no], use_container_width=True)

    elif menu == "📋 Senarai Laporan":
        st.markdown('<div class="card"><h3>📋 Senarai Laporan BLPP</h3></div>', unsafe_allow_html=True)
        if user["peranan"]=="PPD":
            my = [r for r in reports if is_report_for_ppd_strict(r, user["ppd"])]
        elif user["peranan"]=="KP":
            my = [r for r in reports if r.get("id_kp")==user["id"] or r.get("no_pusat")==user.get("no_pusat","")]
        else:
            my = reports
        if my:
            df = pd.DataFrame(my)
            st.dataframe(df, use_container_width=True)
            # Quick view
            sel = st.selectbox("Pilih Laporan untuk Lihat/Cetak", [f"{r['id']} - {r['no_pusat']} - {r['tarikh']}" for r in my])
            if sel:
                rid = sel.split(" - ")[0]
                rep = next((r for r in my if r["id"]==rid), None)
                if rep:
                    html = jana_blpp_html(rep, LOGO_B64)
                    st.components.v1.html(html, height=600, scrolling=True)
                    st.download_button("📥 Download HTML", html, file_name=f"{rid}.html")
        else:
            st.info("Belum ada laporan")

    elif menu == "🖨️ Cetak Borang":
        st.markdown('<div class="card"><h3>🖨️ Cetak Semula Borang</h3></div>', unsafe_allow_html=True)
        st.write("Guna menu Senarai Laporan untuk cetak semula. Logo LP rasmi akan tertera.")

    elif menu == "📊 Statistik JPN":
        st.markdown('<div class="card"><h3>📊 Dashboard Statistik</h3></div>', unsafe_allow_html=True)
        c1,c2,c3,c4 = st.columns(4)
        with c1: st.metric("Total Laporan", len(reports))
        with c2: st.metric("Tiada Kes", len([r for r in reports if r.get("tiada_kes")]))
        with c3: st.metric("Ada Kes", len([r for r in reports if not r.get("tiada_kes")]))
        with c4: st.metric("Hari Ini", len([r for r in reports if r.get("tarikh")==str(date.today())]))
        if reports:
            df = pd.DataFrame(reports)
            st.bar_chart(df.groupby("ppd").size())
            # PPD breakdown
            stats=[]
            for ppd in PPD_LIST:
                cnt = len([r for r in reports if is_report_for_ppd_strict(r, ppd)])
                stats.append({"PPD":ppd,"Laporan":cnt})
            st.dataframe(pd.DataFrame(stats), use_container_width=True)

    else:
        st.markdown('<div class="card"><h3>⚙️ Pengurusan Pengguna (JPN)</h3></div>', unsafe_allow_html=True)
        if user["peranan"]=="ADMIN_JPN":
            users = load_users()
            dfu = pd.DataFrame([u for u in users if u["peranan"]=="KP"])
            st.dataframe(dfu, use_container_width=True, height=400)
            st.download_button("📥 Download Senarai KP", dfu.to_csv(index=False).encode(), file_name="KP_504.csv")
        else:
            st.info("Akses JPN sahaja")

st.markdown('<div style="text-align:center;font-size:10px;color:#94A3B8;margin-top:30px;">BLPP v50 PRO • Lembaga Peperiksaan Malaysia • JPN Selangor 504 Pusat • Integrated SPM + Amali Sains • © 2026</div>', unsafe_allow_html=True)
