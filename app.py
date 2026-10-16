# BLPP JPN Selangor - v51 FINAL UPGRADE - 4 Permintaan Akashah
# 1. KP boleh cari semula laporan dia untuk cetak
# 2. Login pusat tiada dalam data -> keluar pernyataan
# 3. PPD download zip boleh pilih tarikh + peperiksaan
# 4. JPN & LP download zip tambah pilihan peperiksaan + tarikh

import streamlit as st
import pandas as pd
import json, os, base64, zipfile
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

st.set_page_config(page_title="BLPP JPN v51 - LP Official", layout="wide", page_icon="📋", initial_sidebar_state="expanded")

def get_logo_base64():
    for p in ["/mnt/data/logo_lp_transparent.png", "/mnt/data/logo_lp.png", "logo_lp.png"]:
        if os.path.exists(p):
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode()
    return None

LOGO_B64 = get_logo_base64()

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=Plus+Jakarta+Sans:wght@700;800&display=swap');
* {{ font-family: 'Inter', sans-serif; }}
#MainMenu, header, footer, [data-testid="stToolbar"], [data-testid="stHeader"] {{visibility: hidden; display: none;}}
.stApp {{ background: linear-gradient(180deg, #F8FAFF 0%, #EEF2FF 100%) !important; }}
html, body, .stApp {{ color-scheme: light !important; }}
div[data-testid="stSelectbox"] div[data-baseweb="select"],
div[data-testid="stSelectbox"] div[data-baseweb="select"] *,
div[data-testid="stTextInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-testid="stDateInput"] input,
input, textarea, select {{
    color: #000000 !important; background-color: #FFFFFF !important; -webkit-text-fill-color: #000000 !important;
}}
div[data-baseweb="popover"], div[data-baseweb="menu"], ul[data-baseweb="menu"], li[data-baseweb="option"] {{
    background-color: #FFFFFF !important; color: #000000 !important;
}}
.header-pro {{
    background: linear-gradient(135deg, #07112A 0%, #0A1931 20%, #0D47A1 60%, #1565C0 100%);
    padding:26px 28px; border-radius:20px; margin-bottom:20px; text-align:center;
    box-shadow:0 14px 40px rgba(13,71,161,0.3); border:1px solid rgba(255,255,255,0.12);
    position: relative; overflow: hidden;
}}
.header-pro img {{ height:58px; background:white; padding:6px 18px; border-radius:12px; margin-bottom:14px; box-shadow:0 4px 16px rgba(0,0,0,0.2); }}
.header-pro h1 {{ color:white; margin:0; font-family:'Plus Jakarta Sans'; font-size:24px; font-weight:800; position:relative; }}
.header-pro h2 {{ color:#FFECB3; margin:8px 0 0 0; font-size:12px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; position:relative; }}
.header-pro h3 {{ color:white; margin:14px 0 0 0; font-size:14px; font-weight:700; background:rgba(255,255,255,0.18); display:inline-block; padding:8px 22px; border-radius:24px; position:relative; }}
.header-pro .ver {{ position:absolute; top:14px; right:18px; background:#FFD54F; color:#0A1931; font-size:10px; font-weight:800; padding:4px 10px; border-radius:20px; }}
.card {{ background:white; border-radius:16px; padding:20px; border:1px solid #E8EEFF; box-shadow:0 6px 20px rgba(0,0,0,0.06); margin:14px 0; }}
.alert-pusat-tiada {{ background:#FEF2F2; border:2px solid #FCA5A5; color:#991B1B; padding:16px; border-radius:12px; font-weight:700; text-align:center; }}
div[data-testid="stButton"] button {{ border-radius:12px !important; font-weight:700 !important; }}
</style>
""", unsafe_allow_html=True)

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

@st.cache_data
def load_spm_data():
    path = "/mnt/data/data_v1.1_3_1.xlsx"
    if not os.path.exists(path): 
        # cuba cari di local
        if os.path.exists("data_v1.1_3_1.xlsx"):
            path = "data_v1.1_3_1.xlsx"
        else:
            return None, None, None
    try:
        pusat = pd.read_excel(path, sheet_name="selenggara_pusat")
        mp = pd.read_excel(path, sheet_name="Data_Lengkap_4Pakej")
        bilik = pd.read_excel(path, sheet_name="Senarai_Bilik_Kebal")
        return pusat, mp, bilik
    except Exception as e:
        return None,None,None

@st.cache_data
def load_amali_data():
    path = "/mnt/data/data_pusat_amali_sains_baru_1.xlsx"
    if not os.path.exists(path):
        if os.path.exists("data_pusat_amali_sains_baru_1.xlsx"):
            path="data_pusat_amali_sains_baru_1.xlsx"
        else:
            return None,None
    try:
        ringkas = pd.read_excel(path, sheet_name="Ringkasan_Makmal")
        detail = pd.read_excel(path, sheet_name="Data_Pusat_Amali_Sains")
        return ringkas, detail
    except:
        return None,None

# ===== FUNGSI BARU: SEMAK PUSAT ADA DALAM DATA ATAU TIDAK =====
def check_pusat_exists(no_pusat):
    pusat_df,_,_ = load_spm_data()
    if pusat_df is None:
        return True  # jika data tak load, jangan block
    if not no_pusat:
        return True
    no_pusat = str(no_pusat).strip().upper()
    # cek exact match
    exists = pusat_df["No_Pusat"].astype(str).str.upper().str.strip().eq(no_pusat).any()
    return exists

def get_pusat_info(no_pusat):
    pusat_df,_,_ = load_spm_data()
    if pusat_df is None or not no_pusat:
        return None
    row = pusat_df[pusat_df["No_Pusat"].astype(str).str.upper()==str(no_pusat).upper()]
    if not row.empty:
        return row.iloc[0].to_dict()
    return None

def load_users():
    if os.path.exists("users.json"):
        with open("users.json","r") as f: return json.load(f)
    pusat,_,_ = load_spm_data()
    users=[]
    if pusat is not None:
        for _,row in pusat.iterrows():
            no = str(row.get("No_Pusat","")).strip()
            if not no or no=="nan": continue
            kod = no[:2] if len(no)>=2 else "BA"
            ppd = KOD_PPD_MAPPING.get(kod,"Klang")
            users.append({
                "id": no, "nama": f"KP {no}", "peranan":"KP","ppd":ppd,
                "no_pusat":no,"nama_pusat":str(row.get("Nama_Pusat","")),"password":no.lower()
            })
    users.extend([
        {"id":"PPD01","nama":"PPD Petaling Utama","peranan":"PPD","ppd":"Petaling Utama","no_pusat":"","nama_pusat":"","password":"ppd01"},
        {"id":"JPN001","nama":"JPN Selangor","peranan":"ADMIN_JPN","ppd":"","no_pusat":"","nama_pusat":"","password":"jpn123"},
        {"id":"LP001","nama":"Lembaga Peperiksaan","peranan":"LP","ppd":"","no_pusat":"","nama_pusat":"","password":"lp123"},
        {"id":"BH562","nama":"KP BH562 (Contoh)","peranan":"KP","ppd":"Petaling Perdana","no_pusat":"BH562","nama_pusat":"SMK SEKSYEN 19","password":"bh562"},
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

def render_header(user=None):
    logo_html = f'<img src="data:image/png;base64,{LOGO_B64}" alt="LP Logo">' if LOGO_B64 else '📋'
    role_text = f"{user['peranan']} - {user.get('ppd','')} {user.get('no_pusat','')}" if user else "SISTEM BLPP RASMI LP/Am 90 Pin. 1/2023 - v51"
    st.markdown(f"""
    <div class="header-pro">
        <div class="ver">v51 FINAL - 4 FUNGSI BARU</div>
        {logo_html}
        <h1>Sistem Pengurusan Pusat Peperiksaan</h1>
        <h2>Borang Laporan Penyelewengan / Bencana (BLPP) • Lembaga Peperiksaan Malaysia</h2>
        <h3>{role_text}</h3>
    </div>
    """, unsafe_allow_html=True)

def jana_blpp_html(report, logo_b64):
    kes_rows=""
    if report.get("tiada_kes"):
        kes_rows="<tr><td colspan='5' style='text-align:center;padding:20px;'>TIADA KES / TIADA PENYELEWENGAN</td></tr>"
    else:
        for i,k in enumerate(report.get("kes_data",[]),1):
            if not k.get('jenis'): continue
            kes_rows+=f"<tr><td>{i}</td><td>{k.get('jenis','')}</td><td>{k.get('kertas','')}</td><td>{k.get('bil_calon','')}</td><td>{k.get('catatan','')}</td></tr>"
        if not kes_rows:
            kes_rows="<tr><td colspan='5' style='text-align:center;'>TIADA BUTIRAN KES</td></tr>"
    logo_img = f'<img src="data:image/png;base64,{logo_b64}" style="height:55px;">' if logo_b64 else ''
    html=f"""
    <html><head><style>
    body {{ font-family: Arial; font-size:11px; color:#000; }}
    .kop {{ text-align:center; border-bottom:3px double #000; padding-bottom:10px; margin-bottom:12px; }}
    table {{ width:100%; border-collapse:collapse; }} th, td {{ border:1px solid #000; padding:6px; font-size:10px; }} th {{ background:#E8EEFF; }}
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

def create_bulk_zip(filtered_reports, logo_b64):
    """Buat ZIP mengandungi semua HTML BLPP"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for r in filtered_reports:
            html = jana_blpp_html(r, logo_b64)
            fname = f"{r.get('tarikh','')}_{r.get('no_pusat','')}_{r.get('peperiksaan','')}_{r['id']}.html"
            # bersihkan nama file
            fname = "".join(c for c in fname if c.isalnum() or c in ('-','_','.')).replace('__','_')
            zf.writestr(fname, html)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# ========== LOGIN DENGAN SEMAKAN PUSAT TIADA DALAM DATA [FUNGSI 2] ==========
if "user" not in st.session_state:
    st.session_state.user=None

if not st.session_state.user:
    render_header()
    col1,col2,col3 = st.columns([1,1.3,1])
    with col2:
        st.markdown('<div class="card"><h3>🔐 Log Masuk Sistem v51</h3><p style="font-size:12px;color:#64748B;">ID KP = No Pusat (contoh BH562). Jika pusat tiada dalam data, sistem akan maklumkan.</p></div>', unsafe_allow_html=True)
        with st.form("login"):
            id_input = st.text_input("ID Pengguna / No Pusat", placeholder="Contoh: BH562")
            pass_input = st.text_input("Kata Laluan", type="password", placeholder="bh562")
            peranan = st.selectbox("Peranan", ["KP","PPD","ADMIN_JPN","LP"])
            login_btn = st.form_submit_button("Masuk Sistem →", use_container_width=True)
            if login_btn:
                # SEMAKAN 2: PUSAT TIADA DALAM DATA
                if peranan=="KP":
                    no_check = id_input.strip().upper()
                    if no_check and not check_pusat_exists(no_check):
                        st.markdown(f"""
                        <div class="alert-pusat-tiada">
                        ⚠️ PUSAT TIADA DALAM DATA<br><br>
                        No Pusat <b>{no_check}</b> tidak wujud dalam data rasmi JPN Selangor 2026 (504 Pusat).<br>
                        Sila semak No Pusat atau hubungi JPN Selangor / LP untuk pendaftaran.<br><br>
                        Kod yang ada: BA, BB, BC, BD, BE, BF, BG, BH, BJ, BK
                        </div>
                        """, unsafe_allow_html=True)
                        st.stop()
                
                users = load_users()
                found=None
                for u in users:
                    if u["id"].lower()==id_input.lower() and u["password"]==pass_input and u["peranan"]==peranan:
                        found=u; break
                    if u.get("no_pusat","").lower()==id_input.lower() and u["password"]==pass_input:
                        found=u; break
                if found:
                    st.session_state.user=found
                    st.rerun()
                else:
                    # Jika KP tapi ID betul format pusat tapi password salah, masih check data
                    if peranan=="KP" and check_pusat_exists(id_input):
                        st.error(f"ID Pusat {id_input} WUJUD dalam data, tapi Password / Peranan salah. Cuba password: {id_input.lower()}")
                    else:
                        st.error("ID / Password / Peranan salah. Hubungi JPN.")
    st.stop()

# ========== DASHBOARD ==========
user = st.session_state.user
render_header(user)
col_menu, col_content = st.columns([1,3.2])

with col_menu:
    st.markdown(f'<div class="card" style="text-align:center;"><b>{user["nama"]}</b><br><span style="font-size:12px;color:#64748B;">{user["id"]} • {user["peranan"]}</span><br><span style="font-size:11px;background:#EEF2FF;padding:4px 10px;border-radius:20px;">{user.get("ppd","")}</span></div>', unsafe_allow_html=True)
    if user["peranan"]=="KP":
        menu = st.radio("Menu KP", ["📝 Isi Borang BLPP", "🔍 Cari & Cetak Semula Laporan Saya [BARU v51]", "🏫 Data Pusat Saya", "📋 Senarai Laporan"], label_visibility="collapsed")
    elif user["peranan"]=="PPD":
        menu = st.radio("Menu PPD", ["📋 Senarai Laporan PPD", "📦 Download Pukal ZIP (Tarikh + Peperiksaan) [BARU v51]", "🏫 Data PPD", "📊 Statistik"], label_visibility="collapsed")
    elif user["peranan"] in ["ADMIN_JPN","LP"]:
        menu = st.radio("Menu JPN/LP", ["📋 Senarai Semua Laporan", "📦 Download ZIP Pukal (JPN/LP - Pilih Peperiksaan + Tarikh) [BARU v51]", "🏫 Data 504 Pusat & Amali", "📊 Dashboard JPN", "⚙️ Pengurusan"], label_visibility="collapsed")
    else:
        menu = st.radio("Menu", ["📋 Senarai Laporan", "📦 Download ZIP", "📊 Statistik"], label_visibility="collapsed")

    if st.button("Keluar Log", use_container_width=True):
        st.session_state.user=None
        st.rerun()

reports = load_reports()

with col_content:
    # ===================== KP: FUNGSI 1 - CARI SEMULA LAPORAN UNTUK CETAK =====================
    if user["peranan"]=="KP" and "Cari & Cetak" in menu:
        st.markdown('<div class="card"><h3>🔍 [BARU v51] Cari Semula Laporan Saya Untuk Dicetak</h3><p style="font-size:12px;">Fungsi 1 yang awak minta - KP boleh cari laporan lama dia dan cetak semula.</p></div>', unsafe_allow_html=True)
        my_reports = [r for r in reports if r.get("id_kp")==user["id"] or r.get("no_pusat")==user.get("no_pusat","")]
        
        if not my_reports:
            st.info(f"Belum ada laporan untuk {user.get('no_pusat','')}. Sila isi borang dulu.")
        else:
            c1,c2,c3 = st.columns(3)
            with c1:
                tarikh_filter = st.date_input("Tapis Tarikh", value=None, key="kp_tarikh")
            with c2:
                pep_filter = st.selectbox("Tapis Peperiksaan", ["Semua","SPM","SPM Ulangan","STAM","STPM"], key="kp_pep")
            with c3:
                search_id = st.text_input("Cari ID Laporan", placeholder="BLPP-...")

            filtered = my_reports
            if tarikh_filter:
                filtered = [r for r in filtered if r.get("tarikh")==str(tarikh_filter)]
            if pep_filter!="Semua":
                filtered = [r for r in filtered if r.get("peperiksaan")==pep_filter]
            if search_id:
                filtered = [r for r in filtered if search_id.lower() in r.get("id","").lower()]

            st.markdown(f"**{len(filtered)} laporan dijumpai** untuk {user.get('no_pusat','')}")
            
            if filtered:
                df = pd.DataFrame([{"ID":r["id"],"Tarikh":r.get("tarikh"),"Peperiksaan":r.get("peperiksaan"),"Status":"TIADA KES" if r.get("tiada_kes") else f"{len(r.get('kes_data',[]))} KES","Masa Hantar":r.get("tarikh_submit")} for r in filtered])
                st.dataframe(df, use_container_width=True)

                st.markdown("---")
                sel = st.selectbox("📌 Pilih Laporan Untuk Cetak Semula", [f"{r['id']} | {r.get('tarikh')} | {r.get('peperiksaan')} | {'TIADA KES' if r.get('tiada_kes') else 'ADA KES'}" for r in filtered])
                if sel:
                    rid = sel.split(" | ")[0]
                    rep = next((r for r in filtered if r["id"]==rid), None)
                    if rep:
                        html = jana_blpp_html(rep, LOGO_B64)
                        st.components.v1.html(html, height=650, scrolling=True)
                        colA,colB = st.columns(2)
                        with colA:
                            st.download_button("🖨️ Download & Cetak (HTML)", html, file_name=f"{rid}_CETAK_SEMULA.html", mime="text/html", use_container_width=True)
                        with colB:
                            # Buat PDF via download html print
                            st.info("Tip: Buka file HTML > Ctrl+P > Save as PDF untuk dapat logo LP rasmi")

    elif user["peranan"]=="KP" and "Isi Borang" in menu:
        st.markdown('<div class="card"><h3>📝 Borang BLPP - LP/Am 90</h3></div>', unsafe_allow_html=True)
        with st.form("blpp_form"):
            c1,c2,c3 = st.columns(3)
            with c1:
                tarikh = st.date_input("Tarikh Peperiksaan", value=date.today())
                no_pusat = st.text_input("No Pusat", value=user.get("no_pusat",""))
            with c2:
                nama_pusat = st.text_input("Nama Pusat", value=user.get("nama_pusat",""))
                peperiksaan = st.selectbox("Peperiksaan", ["SPM","SPM Ulangan","STAM","STPM","SPM Amali Sains"])
            with c3:
                tahun = st.text_input("Tahun", value="2026")
                ppd = st.text_input("PPD", value=user.get("ppd",""))
            st.markdown("---")
            tiada_kes = st.checkbox("☑ TIADA KES / TIADA PENYELEWENGAN BERLAKU")
            kes_data=[]
            if not tiada_kes:
                for i in range(3):
                    cols = st.columns([1.5,1,0.8,2])
                    with cols[0]: jenis = st.selectbox(f"Jenis Kes {i+1}", ["","Kertas Bocor","Salah Edar Kertas","Calon Meniru","Kecemasan Bencana","Amali Sains - Ralat Alat","Lain-lain"], key=f"j{i}")
                    with cols[1]: kertas = st.text_input(f"Kertas {i+1}", key=f"k{i}")
                    with cols[2]: bil = st.number_input(f"Bil Calon {i+1}", min_value=0, key=f"b{i}")
                    with cols[3]: catatan = st.text_input(f"Catatan {i+1}", key=f"c{i}")
                    if jenis: kes_data.append({"jenis":jenis,"kertas":kertas,"bil_calon":bil,"catatan":catatan})
            nama_kp = st.text_input("Nama Ketua Pengawas", value=user.get("nama",""))
            telefon = st.text_input("No Telefon KP")
            hantar = st.form_submit_button("✅ Hantar Laporan Rasmi", use_container_width=True)
            if hantar:
                # Semak lagi pusat ada dalam data
                if not check_pusat_exists(no_pusat):
                    st.error(f"⚠️ Pusat {no_pusat} TIADA dalam data rasmi. Laporan tidak disimpan. Hubungi JPN.")
                else:
                    new_report = {
                        "id": f"BLPP-{no_pusat}-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                        "tarikh": str(tarikh), "tarikh_submit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "peperiksaan": peperiksaan, "tahun": tahun,
                        "no_pusat": no_pusat, "nama_pusat": nama_pusat,
                        "nama_kp": nama_kp, "ppd": ppd, "id_kp": user["id"],
                        "kes_data": kes_data, "tiada_kes": tiada_kes, "telefon_kp": telefon
                    }
                    reports.append(new_report)
                    save_reports(reports)
                    st.success(f"✅ Laporan {new_report['id']} berjaya dihantar!")
                    html = jana_blpp_html(new_report, LOGO_B64)
                    st.components.v1.html(html, height=600, scrolling=True)
                    st.download_button("📥 Download Borang", html, file_name=f"{new_report['id']}.html")

    # ===================== PPD: FUNGSI 3 - DOWNLOAD ZIP PILIH TARIKH + PEPERIKSAAN =====================
    elif "Download Pukal ZIP" in menu:
        if user["peranan"]=="PPD":
            st.markdown('<div class="card"><h3>📦 [BARU v51] Download Pukal ZIP - PPD (Pilih Tarikh + Peperiksaan)</h3><p style="font-size:12px;">Fungsi 3 yang awak minta - PPD boleh tapis ikut tarikh dan jenis peperiksaan sebelum download ZIP.</p></div>', unsafe_allow_html=True)
            # Filter reports ikut PPD user
            my_reports = [r for r in reports if is_report_for_ppd_strict(r, user["ppd"])]
            if not my_reports:
                st.info(f"Belum ada laporan untuk PPD {user['ppd']}")
            else:
                c1,c2,c3 = st.columns(3)
                with c1:
                    tarikh_mula = st.date_input("Tarikh Mula", value=None, key="ppd_mula")
                    tarikh_tamat = st.date_input("Tarikh Tamat", value=None, key="ppd_tamat")
                with c2:
                    pep_list = ["Semua"] + sorted(list(set([r.get("peperiksaan","SPM") for r in my_reports])))
                    pep_pilih = st.multiselect("Pilih Peperiksaan", pep_list, default=["Semua"], key="ppd_pep")
                with c3:
                    st.metric("Total Laporan PPD", len(my_reports))
                    # Filter logic
                    filtered = my_reports
                    if tarikh_mula:
                        filtered = [r for r in filtered if r.get("tarikh","") >= str(tarikh_mula)]
                    if tarikh_tamat:
                        filtered = [r for r in filtered if r.get("tarikh","") <= str(tarikh_tamat)]
                    if pep_pilih and "Semua" not in pep_pilih:
                        filtered = [r for r in filtered if r.get("peperiksaan") in pep_pilih]
                    st.metric("Akan Di-Download", len(filtered))

                if filtered:
                    st.dataframe(pd.DataFrame([{"No Pusat":r["no_pusat"],"Tarikh":r["tarikh"],"Peperiksaan":r["peperiksaan"],"Nama Pusat":r["nama_pusat"]} for r in filtered]), use_container_width=True)
                    if st.button(f"📦 Jana & Download ZIP ({len(filtered)} Borang)", type="primary", use_container_width=True):
                        zip_bytes = create_bulk_zip(filtered, LOGO_B64)
                        st.download_button("⬇️ Download ZIP Sekarang", zip_bytes, file_name=f"BLPP_PPD_{user['ppd']}_{date.today()}_{len(filtered)}_borang.zip", mime="application/zip", use_container_width=True)
                        st.success(f"{len(filtered)} borang berjaya di-ZIP!")

        # ===================== JPN & LP: FUNGSI 4 - DOWNLOAD ZIP PILIH PEPERIKSAAN + TARIKH =====================
        else:
            st.markdown('<div class="card"><h3>📦 [BARU v51] Download Pukal ZIP - JPN / LP (Pilih Peperiksaan + Tarikh + PPD)</h3><p style="font-size:12px;">Fungsi 4 yang awak minta - JPN & LP boleh pilih PPD, tarikh, dan peperiksaan.</p></div>', unsafe_allow_html=True)
            if not reports:
                st.info("Belum ada laporan")
            else:
                c1,c2,c3,c4 = st.columns(4)
                with c1:
                    ppd_pilih = st.multiselect("Pilih PPD (JPN/LP)", ["Semua"]+PPD_LIST, default=["Semua"])
                with c2:
                    tarikh_mula = st.date_input("Tarikh Mula", value=None, key="jpn_mula")
                    tarikh_tamat = st.date_input("Tarikh Tamat", value=None, key="jpn_tamat")
                with c3:
                    pep_all = sorted(list(set([r.get("peperiksaan","SPM") for r in reports])))
                    pep_pilih = st.multiselect("Pilih Peperiksaan", ["Semua"]+pep_all, default=["Semua"], key="jpn_pep")
                with c4:
                    st.metric("Total Semua Laporan", len(reports))

                filtered = reports
                if ppd_pilih and "Semua" not in ppd_pilih:
                    filtered = [r for r in filtered if any(is_report_for_ppd_strict(r, ppd) for ppd in ppd_pilih)]
                if tarikh_mula:
                    filtered = [r for r in filtered if r.get("tarikh","") >= str(tarikh_mula)]
                if tarikh_tamat:
                    filtered = [r for r in filtered if r.get("tarikh","") <= str(tarikh_tamat)]
                if pep_pilih and "Semua" not in pep_pilih:
                    filtered = [r for r in filtered if r.get("peperiksaan") in pep_pilih]

                st.markdown(f"### {len(filtered)} laporan akan di-ZIP selepas tapisan")
                if filtered:
                    st.dataframe(pd.DataFrame([{"No Pusat":r["no_pusat"],"PPD":r.get("ppd"),"Tarikh":r["tarikh"],"Peperiksaan":r["peperiksaan"],"Status":"TIADA KES" if r.get("tiada_kes") else "ADA KES"} for r in filtered]), use_container_width=True, height=400)
                    
                    colZ1,colZ2 = st.columns(2)
                    with colZ1:
                        if st.button(f"📦 Jana ZIP {len(filtered)} Borang (JPN/LP)", type="primary", use_container_width=True):
                            zip_bytes = create_bulk_zip(filtered, LOGO_B64)
                            st.session_state["zip_jpn"] = zip_bytes
                            st.success("ZIP berjaya dijana! Klik download di sebelah")
                    with colZ2:
                        if "zip_jpn" in st.session_state:
                            st.download_button("⬇️ Download ZIP JPN/LP Sekarang", st.session_state["zip_jpn"], file_name=f"BLPP_JPN_SELANGOR_{date.today()}_{len(filtered)}_borang_PPD_{'_'.join(ppd_pilih[:2])}_{'_'.join(pep_pilih[:2])}.zip", mime="application/zip", use_container_width=True)

    elif "Data" in menu or "Senarai" in menu or "Statistik" in menu:
        # Reuse simple views
        if "Senarai" in menu:
            if user["peranan"]=="PPD":
                my = [r for r in reports if is_report_for_ppd_strict(r, user["ppd"])]
            elif user["peranan"]=="KP":
                my = [r for r in reports if r.get("id_kp")==user["id"] or r.get("no_pusat")==user.get("no_pusat","")]
            else:
                my = reports
            st.dataframe(pd.DataFrame(my), use_container_width=True)
        elif "Data" in menu:
            pusat_df, mp_df, _ = load_spm_data()
            ringkas_df, detail_df = load_amali_data()
            if pusat_df is not None:
                if user["peranan"]=="PPD":
                    kods = get_kods_for_ppd(user["ppd"])
                    pusat_df = pusat_df[pusat_df["No_Pusat"].astype(str).str[:2].isin(kods)]
                st.dataframe(pusat_df, use_container_width=True, height=500)
            if ringkas_df is not None:
                st.markdown("**Data Amali Sains**")
                st.dataframe(ringkas_df.head(200), use_container_width=True)
        else:
            c1,c2,c3 = st.columns(3)
            with c1: st.metric("Total Laporan", len(reports))
            with c2: st.metric("Tiada Kes", len([r for r in reports if r.get("tiada_kes")]))
            with c3: st.metric("Ada Kes", len([r for r in reports if not r.get("tiada_kes")]))
            if reports:
                df = pd.DataFrame(reports)
                st.bar_chart(df.groupby("ppd").size())

    else:
        st.info("Pilih menu di sebelah kiri")

st.markdown('<div style="text-align:center;font-size:10px;color:#94A3B8;margin-top:30px;">BLPP v51 FINAL • 4 Fungsi Baru Akashah • Lembaga Peperiksaan Malaysia • JPN Selangor 504 Pusat • © 2026</div>', unsafe_allow_html=True)
