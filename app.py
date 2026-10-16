# App-50-v41-FINAL-QA-FIX - Production 504 Pusat JPN Selangor - KP/PPD/JPN/LP - Dark Mode Fix + Pep Filter + KP Cetak Semula
# Based on v40 perfect - All functions retained, debug cleaned, garis lurus, tak senggang, logo LP
import streamlit as st
import time
import pandas as pd
import json, os
from datetime import datetime, date
from io import BytesIO
import base64
from PIL import Image
import io

# ========== GOOGLE SHEETS INTEGRATION (OPTIONAL) ==========
# Jika ada credentials, akan guna Google Sheets. Jika tiada, guna JSON local (demo)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except:
    GSHEETS_AVAILABLE = False

def get_gsheets_client():
    if not GSHEETS_AVAILABLE:
        return None
    try:
        # Guna st.secrets untuk credentials
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            )
            client = gspread.authorize(creds)
            return client
    except Exception as e:
        st.warning(f"Google Sheets tidak dapat disambung: {e}, guna mode local JSON")
    return None

@st.cache_data(ttl=300, show_spinner=False)
def load_reports_from_gsheets_cached(sheet_id_cache):
    client = get_gsheets_client()
    if not client:
        return None
    try:
        sh = client.open_by_key(sheet_id_cache)
        ws = sh.worksheet("BLPP_Reports")
        records = ws.get_all_records()
        return records
    except Exception as e:
        if "429" in str(e) or "Quota" in str(e):
            return "QUOTA_429"
        return None

def load_reports_from_gsheets():
    client = get_gsheets_client()
    if not client:
        return None
    try:
        DEFAULT_SHEET_ID = "1jT3FknnnC3iCzn_xrJNWK-XfLY3ljIZwzvyVYdx3rUc"
        sheet_id = DEFAULT_SHEET_ID
        try:
            if "gsheets_id" in st.secrets:
                sheet_id = st.secrets.get("gsheets_id", sheet_id)
            elif "gsheet_id" in st.secrets:
                sheet_id = st.secrets.get("gsheet_id", sheet_id)
            elif "gsheets" in st.secrets:
                g = st.secrets["gsheets"]
                sheet_id = g.get("sheet_id", g.get("gsheets_id", sheet_id)) if hasattr(g, 'get') else g["sheet_id"]
            if not sheet_id and "sheet_id" in st.secrets:
                sheet_id = st.secrets.get("sheet_id", DEFAULT_SHEET_ID)
        except:
            pass
        if not sheet_id:
            sheet_id = DEFAULT_SHEET_ID
        records_raw = load_reports_from_gsheets_cached(sheet_id)
        if records_raw == "QUOTA_429" or records_raw is None:
            return None
        records = records_raw
        reports = []
        for r in records:
            try:
                kes_data = []
                try:
                    import json as js
                    if r.get("kes_data_json"):
                        kes_data = js.loads(r.get("kes_data_json"))
                except:
                    kes_data = []
                reports.append({
                    "id": r.get("id",""),
                    "tarikh": r.get("tarikh",""),
                    "tarikh_submit": r.get("tarikh_submit",""),
                    "peperiksaan": r.get("peperiksaan",""),
                    "tahun": r.get("tahun",""),
                    "no_pusat": r.get("no_pusat",""),
                    "nama_pusat": r.get("nama_pusat",""),
                    "nama_kp": r.get("nama_kp",""),
                    "ppd": r.get("ppd",""),
                    "id_kp": r.get("id_kp",""),
                    "kes_data": kes_data,
                    "pelapor": r.get("pelapor",""),
                    "telefon_kp": r.get("telefon_kp","")
                })
            except:
                continue
        return reports
    except Exception as e:
        return None

def save_to_gsheets(report):
    client = get_gsheets_client()
    if not client:
        return False
    try:
        DEFAULT_SHEET_ID = "1jT3FknnnC3iCzn_xrJNWK-XfLY3ljIZwzvyVYdx3rUc"
        sheet_id = DEFAULT_SHEET_ID
        try:
            if "gsheets_id" in st.secrets:
                sheet_id = st.secrets.get("gsheets_id", sheet_id)
            elif "gsheet_id" in st.secrets:
                sheet_id = st.secrets.get("gsheet_id", sheet_id)
            elif "gsheets" in st.secrets:
                g = st.secrets["gsheets"]
                sheet_id = g.get("sheet_id", g.get("gsheets_id", sheet_id)) if hasattr(g, 'get') else g["sheet_id"]
            if not sheet_id and "sheet_id" in st.secrets:
                sheet_id = st.secrets.get("sheet_id", DEFAULT_SHEET_ID)
        except:
            pass
        if not sheet_id:
            sheet_id = DEFAULT_SHEET_ID
        for attempt in range(3):
            try:
                sh = client.open_by_key(sheet_id)
                ws = sh.worksheet("BLPP_Reports")
                row = [
                    report["id"],
                    report["tarikh"],
                    report["tarikh_submit"],
                    report["peperiksaan"],
                    report["tahun"],
                    report["no_pusat"],
                    report["nama_pusat"],
                    report["nama_kp"],
                    report["ppd"],
                    report["id_kp"],
                    json.dumps(report["kes_data"], ensure_ascii=False),
                    "TIADA KES" if report.get("tiada_kes") else f"{len(report.get('kes_data',[]))} KES",
                    report.get("pelapor",""),
                    report.get("tandatangan_b64","")[:5000],
                    report.get("telefon_kp","")
                ]
                ws.append_row(row)
                try:
                    load_reports_from_gsheets_cached.clear()
                except:
                    pass
                return True
            except Exception as e:
                if "429" in str(e) or "Quota" in str(e):
                    if attempt < 2:
                        time.sleep(2 * (attempt+1))
                        continue
                    else:
                        st.warning("Google Sheets busy (quota). Laporan disimpan local, akan sync auto nanti.")
                        return False
                else:
                    raise e
        return False
    except Exception as e:
        if "429" in str(e) or "Quota" in str(e):
            st.warning("Google Sheets quota penuh, tunggu 1 minit. Data tetap selamat.")
            return False
        st.error(f"Gagal simpan ke Google Sheets: {e}")
        return False


# ========== CONFIG ==========
st.set_page_config(
    page_title="Sistem Pengurusan Pusat Peperiksaan - BLPP LP/Am 90",
    layout="wide",
    page_icon="📋",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap');
* { font-family: 'Inter', sans-serif; }
#MainMenu, header, footer, [data-testid="stToolbar"], [data-testid="stHeader"] {visibility: hidden; display: none;}
.stApp { background: linear-gradient(180deg, #F8FAFF 0%, #EEF2FF 100%) !important; }
div[data-testid="stButton"] button { border-radius:10px !important; font-weight:700 !important; }
div[data-testid="stSelectbox"] div[data-baseweb="select"], div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea, div[data-testid="stDateInput"] input {
    border-radius:10px !important; border:1.5px solid #E2E8F0 !important;
}

/* v36 DARK MODE FIX - Paksa tulisan hitam background putih, tak kira phone dark mode */
html, body, .stApp {
    color-scheme: light !important;
}
div[data-testid="stSelectbox"] div[data-baseweb="select"],
div[data-testid="stSelectbox"] div[data-baseweb="select"] *,
div[data-testid="stTextInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-testid="stDateInput"] input,
input, textarea, select {
    color: #000000 !important;
    background-color: #FFFFFF !important;
    -webkit-text-fill-color: #000000 !important;
}
div[data-baseweb="select"] div,
div[data-baseweb="select"] span,
div[data-baseweb="select"] input {
    color: #000000 !important;
    background: #FFFFFF !important;
}
/* Dropdown menu - pastikan background putih tulisan hitam */
div[data-baseweb="popover"],
div[data-baseweb="menu"],
ul[data-baseweb="menu"],
li[data-baseweb="option"] {
    background-color: #FFFFFF !important;
    color: #000000 !important;
}
li[data-baseweb="option"] * {
    color: #000000 !important;
}
li[data-baseweb="option"]:hover,
li[data-baseweb="option"][aria-selected="true"] {
    background-color: #EEF2FF !important;
    color: #000000 !important;
}
/* Label Peranan */
div[data-testid="stSelectbox"] label,
div[data-testid="stTextInput"] label,
p, label {
    color: #000000 !important;
}

.header-biru {
    background: linear-gradient(135deg, #0A1931 0%, #0D47A1 50%, #1565C0 100%);
    padding:22px 26px; border-radius:16px; margin-bottom:18px; text-align:center;
    box-shadow:0 10px 30px rgba(13,71,161,0.25); border:1px solid rgba(255,255,255,0.1);
}
.header-biru h1 { color:white; margin:0; font-family:'Plus Jakarta Sans'; font-size:22px; font-weight:800; }
.header-biru h2 { color:#FFECB3; margin:6px 0 0 0; font-size:12px; font-weight:700; letter-spacing:1px; }
.header-biru h3 { color:white; margin:10px 0 0 0; font-size:15px; font-weight:700; background:rgba(255,255,255,0.15); display:inline-block; padding:6px 18px; border-radius:20px; }
.card { background:white; border-radius:14px; padding:18px; border:1px solid #EEF2FF; box-shadow:0 4px 16px rgba(0,0,0,0.06); margin:12px 0; }
</style>
""", unsafe_allow_html=True)

# ========== DATA PPD SELANGOR (10 PPD) + KOD PPD ==========
PPD_LIST = [
    "Petaling Utama", "Petaling Perdana", "Gombak", "Hulu Langat", 
    "Klang", "Kuala Langat", "Kuala Selangor", "Hulu Selangor", "Sepang", "Sabak Bernam"
]

KOD_PPD_MAPPING = {
    "BA": "Klang",
    "BB": "Kuala Langat",
    "BC": "Kuala Selangor",
    "BD": "Hulu Langat",
    "BE": "Hulu Selangor",
    "BF": "Sabak Bernam",
    "BG": "Gombak",
    "BH": "Petaling Perdana",
    "BJ": "Sepang",
    "BK": "Petaling Utama"
}


# ===== v29 STRICT PPD FILTER - KASHAH: PPD Klang jangan campur PPD lain - STRICT BY NO PUSAT PREFIX =====
PPD_TO_KODS_STRICT = {
    "Klang": ["BA"],
    "Kuala Langat": ["BB"],
    "Kuala Selangor": ["BC"],
    "Hulu Langat": ["BD"],
    "Hulu Selangor": ["BE"],
    "Sabak Bernam": ["BF"],
    "Gombak": ["BG"],
    "Petaling Perdana": ["BH"],
    "Sepang": ["BJ"],
    "Petaling Utama": ["BK"],
    "Petaling": ["BH", "BK", "BJ"]  # fallback
}

def get_kods_for_ppd(ppd_name):
    """Return list of kod like ['BA'] for Klang, ['BF','BG'] for Gombak"""
    if not ppd_name:
        return []
    ppd_name = str(ppd_name).strip()
    # Direct mapping
    if ppd_name in PPD_TO_KODS_STRICT:
        return PPD_TO_KODS_STRICT[ppd_name]
    # If ppd_name is kod like BA, return [BA]
    if len(ppd_name)==2 and ppd_name.upper() in KOD_PPD_MAPPING:
        return [ppd_name.upper()]
    # Try find by containing name
    for k, v in PPD_TO_KODS_STRICT.items():
        if k.lower() in ppd_name.lower() or ppd_name.lower() in k.lower():
            return v
    return []

def is_report_for_ppd_strict(report, target_ppd):
    """STRICT: Check if report belongs to target PPD by No Pusat prefix ONLY - no mixing"""
    no_pusat = str(report.get("no_pusat","") or "").strip().upper()
    kod_ppd_field = str(report.get("kod_ppd","") or "").strip().upper()
    # Get allowed kods for this PPD
    allowed_kods = get_kods_for_ppd(target_ppd)
    if not allowed_kods:
        return False
    # Check prefix from no_pusat (most reliable - from Excel)
    prefix = no_pusat[:2] if len(no_pusat)>=2 else ""
    if prefix and prefix in allowed_kods:
        return True
    # Check kod_ppd field
    if kod_ppd_field and kod_ppd_field in allowed_kods:
        return True
    return False

def is_report_for_ppd_kod_strict(report, target_kod):
    """STRICT by kod: BA -> only BA reports"""
    target_kod = str(target_kod).strip().upper()
    no_pusat = str(report.get("no_pusat","") or "").strip().upper()
    kod_field = str(report.get("kod_ppd","") or "").strip().upper()
    prefix = no_pusat[:2] if len(no_pusat)>=2 else ""
    return prefix == target_kod or kod_field == target_kod




def get_all_pusat():
    # v24 FINAL: Super flexible - support data_v1.1(3)(1).xlsx macam di Github Kashah
    try:
        import pandas as pd, os, glob
        # Cari semua xlsx yang ada v1.1 atau pusat
        search_patterns = [
            "data_v1.1*.xlsx", "data_v1.1(*).xlsx", "*v1.1*.xlsx",
            "data/*.xlsx", "*.xlsx", "/mnt/data/*.xlsx"
        ]
        candidates = []
        for pat in search_patterns:
            candidates.extend(glob.glob(pat))
            candidates.extend(glob.glob(os.path.join(".", pat)))
        # tambah file spesifik github
        for f in os.listdir("."):
            if f.endswith(".xlsx") and ("v1.1" in f or "pusat" in f.lower()):
                if f not in candidates:
                    candidates.append(f)
        # deduplicate
        candidates = list(set(candidates))
        for p in candidates:
            if not os.path.exists(p):
                continue
            # skip amali
            if "amali" in p.lower():
                continue
            try:
                df = pd.read_excel(p)
                # check ada Kod_PPD dan No_Pusat
                if "Kod_PPD" not in df.columns or "No_Pusat" not in df.columns:
                    continue
                centers=[]
                for _, row in df.iterrows():
                    kod = str(row.get("Kod_PPD","")).strip().upper()
                    if not kod or kod == "NAN":
                        continue
                    centers.append({
                        "Kod_PPD": kod,
                        "No_Pusat": str(row.get("No_Pusat","")).strip(),
                        "Nama_Pusat": str(row.get("Nama_Pusat","")).strip(),
                        "PPD": KOD_PPD_MAPPING.get(kod, str(row.get("PPD", KOD_PPD_MAPPING.get(kod,"")))),
                        "Nama_Sekolah": str(row.get("Nama_Pusat","")).strip(),
                        "Bil_Calon": int(row.get("Bil_Calon_Pusat",0)) if pd.notna(row.get("Bil_Calon_Pusat",0)) else 0
                    })
                if len(centers)>=100:  # asal 400, kita longgarkan untuk test
                    has_bh562 = any(c["No_Pusat"]=="BH562" for c in centers)
                    if not has_bh562:
                        centers.append({"Kod_PPD":"BH","No_Pusat":"BH562","Nama_Pusat":"MAAHAD SAINS AL-QURAN AL-AMIN","PPD":"Petaling Perdana","Nama_Sekolah":"MAAHAD SAINS AL-QURAN AL-AMIN"})
                    print(f"get_all_pusat loaded {len(centers)} from {p}")
                    return centers
            except Exception as e:
                # print(f"get_all_pusat fail {p}: {e}")
                pass
    except Exception as e:
        pass
    # fallback asal - dummy
    centers=[]
    for kod, ppd in KOD_PPD_MAPPING.items():
        for j in range(1, 60):
            if len(centers)>=504:
                break
            centers.append({"Kod_PPD": kod, "No_Pusat": f"{kod}{j:03d}", "Nama_Pusat": f"SMK {ppd} {j}", "PPD": ppd, "Nama_Sekolah": f"SMK {ppd} {j}"})
        if len(centers)>=504:
            break
    if centers:
        centers[0] = {"Kod_PPD":"BH","No_Pusat":"BH562","Nama_Pusat":"MAAHAD SAINS AL-QURAN AL-AMIN","PPD":"Petaling Perdana","Nama_Sekolah":"MAAHAD SAINS AL-QURAN AL-AMIN"}
    return centers

def ensure_centers(centers):
    if not centers or len(centers)==0:
        return get_all_pusat()
    return centers

PPD_TO_KOD = {
    "Klang": "BA",
    "Kuala Langat": "BB",
    "Kuala Selangor": "BC",
    "Hulu Langat": "BD",
    "Hulu Selangor": "BE",
    "Sabak Bernam": "BF",
    "Gombak": "BG",
    "Petaling Perdana": "BH",
    "Sepang": "BJ",
    "Petaling Utama": "BK"
}

PEPERIKSAAN_LIST = [
    "SPM 2026", 
    "SPM Ulangan 2026", 
    "Ujian Amali Sains SPM", 
    "STAM 2026",
    "UJIAN BERTUTUR BAHASA MELAYU",
    "UJIAN BERTUTUR BAHASA INGGERIS",
    "UJIAN MENDENGAR"
]  # FIX KASHAH: Hanya 3 tambah je - Bertutur BM, Bertutur BI, Mendengar - HURUF BESAR

# KASHAH FREEZE - STAM dibekukan di web sebenar
FROZEN_PEPERIKSAAN = ["STAM 2026"]
STAM_ENABLED = False  # True = STAM muncul, False = STAM hilang (freeze)

def get_active_peperiksaan_list():
    if STAM_ENABLED:
        return PEPERIKSAAN_LIST
    else:
        return [p for p in PEPERIKSAAN_LIST if p not in FROZEN_PEPERIKSAAN]

JENIS_PEPERIKSAAN = ["Peperiksaan SPM (504 Pusat)", "Ujian Amali Sains (277 Pusat Makmal)", "Ujian Bertutur & Mendengar"]

def find_excel_file(folder, keywords):
    import os, glob
    if not os.path.exists(folder):
        return None
    for file in os.listdir(folder):
        if file.endswith('.xlsx'):
            lower = file.lower()
            if all(k.lower() in lower for k in keywords):
                return os.path.join(folder, file)
    # fallback cari apa saja xlsx
    files = glob.glob(os.path.join(folder, "*.xlsx"))
    return files[0] if files else None

def load_spm_centers():
    import json, os, glob
    # 1. Cuba JSON
    for p in ["/tmp/blpp_data/spm_centers.json", "data/spm_centers.json"]:
        if os.path.exists(p):
            try:
                with open(p, "r") as f:
                    return json.load(f)
            except:
                pass
    # 2. Cari Excel SPM - super fleksibel untuk nama (3)(1)
    search_folders = ["data", "/tmp/blpp_data", ".", "./data", "/mnt/data"]
    for folder in search_folders:
        if not os.path.exists(folder):
            continue
        for file in os.listdir(folder):
            if not file.endswith('.xlsx'):
                continue
            if "amali" in file.lower():
                continue
            if not ('v1.1' in file or 'spm' in file.lower() or 'pusat' in file.lower()):
                # check any xlsx with Kod_PPD inside
                pass
            path = os.path.join(folder, file)
            try:
                # cuba baca sheet selenggara_pusat dulu
                try:
                    xls = pd.ExcelFile(path)
                    if "selenggara_pusat" in xls.sheet_names:
                        df = pd.read_excel(path, sheet_name="selenggara_pusat")
                    else:
                        df = pd.read_excel(path)
                except:
                    df = pd.read_excel(path)
                if "Kod_PPD" not in df.columns or "No_Pusat" not in df.columns:
                    continue
                centers = []
                for _, row in df.iterrows():
                    kod = str(row.get("Kod_PPD","")).strip()
                    if not kod or kod.upper()=="NAN":
                        continue
                    centers.append({
                        "Kod_PPD": kod.upper(),
                        "No_Pusat": str(row.get("No_Pusat","")).strip(),
                        "Nama_Pusat": str(row.get("Nama_Pusat","")).strip(),
                        "PPD": KOD_PPD_MAPPING.get(kod.upper(), kod),
                        "Bil_Calon": int(row.get("Bil_Calon_Pusat",0)) if pd.notna(row.get("Bil_Calon_Pusat",0)) else 0
                    })
                if centers:
                    # print(f"load_spm_centers loaded {len(centers)} from {path}")
                    return centers
            except Exception as e:
                continue
    return []

def load_amali_centers():
    import json, os
    for p in ["/tmp/blpp_data/amali_centers.json", "data/amali_centers.json"]:
        if os.path.exists(p):
            try:
                with open(p, "r") as f:
                    return json.load(f)
            except:
                pass
    search_folders = ["data", "/tmp/blpp_data", ".", "./data"]
    for folder in search_folders:
        if not os.path.exists(folder):
            continue
        for file in os.listdir(folder):
            if file.endswith('.xlsx') and ('amali' in file.lower()):
                try:
                    path = os.path.join(folder, file)
                    xls = pd.ExcelFile(path)
                    sheet = "Data_Pusat_Amali_Sains" if "Data_Pusat_Amali_Sains" in xls.sheet_names else xls.sheet_names[1]
                    df = pd.read_excel(path, sheet_name=sheet)
                    # Cari kolum yang ada Kod_PPD, No_Pusat, Nama_Sekolah
                    if "Kod_PPD" in df.columns and "No_Pusat" in df.columns:
                        df_unique = df[["Kod_PPD","No_Pusat","Nama_Sekolah"]].drop_duplicates() if "Nama_Sekolah" in df.columns else df[["Kod_PPD","No_Pusat","Nama_Pusat"]].drop_duplicates()
                        centers = []
                        for _, row in df_unique.iterrows():
                            kod = str(row["Kod_PPD"]).strip()
                            nama = str(row.get("Nama_Sekolah", row.get("Nama_Pusat",""))).strip()
                            centers.append({
                                "Kod_PPD": kod,
                                "No_Pusat": str(row["No_Pusat"]).strip(),
                                "Nama_Sekolah": nama,
                                "Nama_Pusat": nama,
                                "PPD": KOD_PPD_MAPPING.get(kod, kod)
                            })
                        if centers:
                            return centers
                except:
                    continue
    return []


# ========== SIMULASI DATABASE (GANTI DENGAN GOOGLE SHEETS NANTI) ==========
DATA_DIR = "/tmp/blpp_data"
os.makedirs(DATA_DIR, exist_ok=True)
USERS_FILE = os.path.join(DATA_DIR, "users.json")
# v27 EMERGENCY RESET - untuk clear users.json lama
def emergency_reset_users():
    import os
    try:
        if os.path.exists(USERS_FILE):
            os.remove(USERS_FILE)
        if os.path.exists(REPORTS_FILE):
            # jangan delete reports, cuma users
            pass
        st.cache_data.clear()
        return True
    except Exception as e:
        return False
REPORTS_FILE = os.path.join(DATA_DIR, "blpp_reports.json")


def find_user_flexible(users, entered_id, entered_password):
    # v26: Support login dengan ID = No Pusat, case-insensitive
    eid = str(entered_id).strip().upper()
    epw = str(entered_password).strip().lower()
    epw_upper = str(entered_password).strip().upper()
    # 1. Cari exact ID match (BA001)
    for u in users:
        uid = str(u.get("id","")).strip().upper()
        pw = str(u.get("password","")).strip().lower()
        pw_alt = str(u.get("password_alt","") or "").strip().lower()
        pw_alt2 = str(u.get("password_alt2","") or "").strip().lower()
        # ID match
        if uid == eid:
            # Password check - lower, upper, or alt
            if pw == epw or pw_alt == epw or pw_alt2 == epw or pw == epw_upper.lower() or str(u.get("password","")).lower() == epw:
                return u
            # Also allow password = no_pusat lower
            if str(u.get("no_pusat","")).strip().lower() == epw:
                return u
    # 2. Cari by no_pusat (kalau user taip BA001 tapi ID dalam file KP001)
    for u in users:
        np = str(u.get("no_pusat","")).strip().upper()
        if np == eid:
            pw = str(u.get("password","")).strip().lower()
            pw_alt = str(u.get("password_alt","") or "").strip().lower()
            pw_alt2 = str(u.get("password_alt2","") or "").strip().lower()
            if pw == epw or pw_alt == epw or pw_alt2 == epw or np.lower() == epw:
                return u
    return None

def load_users():
    # v26 FIX KASHAH: Force regenerate kalau BA001 tak wujud - untuk support ID = No Pusat
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                data = json.load(f)
                kp_count = len([u for u in data if u.get("peranan")=="KP"])
                # Check kalau ada BA001 sebagai ID - kalau tak ada, file lama KP001 style, kena regenerate
                has_ba001 = any(str(u.get("id","")).upper()=="BA001" for u in data)
                has_bk002 = any(str(u.get("id","")).upper()=="BK002" for u in data)
                if kp_count >= 500 and has_ba001 and has_bk002:
                    return data
                else:
                    # File lama - delete untuk regenerate dengan ID = No Pusat
                    try:
                        os.remove(USERS_FILE)
                    except:
                        pass
        except Exception as e:
            try:
                os.remove(USERS_FILE)
            except:
                pass

    # Default users - 504 KP + admin
    users = []
    # Admin JPN
    users.append({"id":"JPN001","nama":"JPN Selangor - Admin","peranan":"ADMIN_JPN","ppd":"JPN","no_pusat":"-","password":"2026","telefon":"","telefon":""})
    users.append({"id":"JPN002","nama":"JPN Selangor 2","peranan":"ADMIN_JPN","ppd":"JPN","no_pusat":"-","password":"jpn123","telefon":""})
    # LP
    users.append({"id":"LP","nama":"Lembaga Peperiksaan","peranan":"LP","ppd":"Semua","no_pusat":"-","password":"2026","telefon":""})
    users.append({"id":"LP001","nama":"Lembaga Peperiksaan","peranan":"LP","ppd":"Semua","no_pusat":"-","password":"2026","telefon":""})
    # 10 PPD - PASSWORD KOD PPD IKUT CADANGAN KASHAH: BA=Klang, BB=Kuala Langat dll - MUDAH INGAT
    # Mapping Kod ke PPD dari KOD_PPD_MAPPING
    PPD_TO_KOD = {
    "Klang": "BA",
    "Kuala Langat": "BB",
    "Kuala Selangor": "BC",
    "Hulu Langat": "BD",
    "Hulu Selangor": "BE",
    "Sabak Bernam": "BF",
    "Gombak": "BG",
    "Petaling Perdana": "BH",
    "Sepang": "BJ",
    "Petaling Utama": "BK"
}  # Klang->BA etc
    for i, ppd in enumerate(PPD_LIST):
        kod = PPD_TO_KOD.get(ppd, f"PPD{i+1:02d}")
        # Password = kod kecil dan besar - mudah: ba, BA, klang, ppd01 semua boleh
        users.append({
            "id": f"PPD{i+1:02d}",
            "nama": f"PPD {ppd} ({kod})",
            "peranan": "PPD",
            "ppd": ppd,
            "no_pusat": "-",
            "password": kod.lower(),  # Password utama = kod kecil: ba, bb, bc...
            "password_alt": kod.upper(),  # Alt = BA, BB
            "kod_ppd": kod,
            "telefon": ""
        })
    # Tambah user login dengan kod terus: BA login sebagai PPD Klang
    for kod, ppd in KOD_PPD_MAPPING.items():
        users.append({
            "id": kod,  # ID = BA, BB, BC...
            "nama": f"PPD {ppd} ({kod}) - Login Kod",
            "peranan": "PPD",
            "ppd": ppd,
            "no_pusat": "-",
            "password": kod.lower(),  # ba
            "password_alt2": kod.upper(),  # BA
            "password_alt3": kod,  # BA
            "kod_ppd": kod,
            "telefon": ""
        })

    # Cuba load SPM centers sebenar untuk dapat Nama Pusat
    spm_centers = []
    try:
        spm_centers = load_spm_centers()
    except:
        spm_centers = []

    if spm_centers and len(spm_centers) >= 10:
        for idx, c in enumerate(spm_centers[:504], 1):
            no_pusat = str(c.get("No_Pusat","")).strip().upper()
            nama_pusat = str(c.get("Nama_Pusat") or c.get("Nama_Sekolah","")).strip()
            ppd = str(c.get("PPD","")).strip()
            # v25 FIX KASHAH: ID = No Pusat terus, bukan KP001 - senang KP ingat
            # Contoh: BA001 login dengan ID BA001 / ba001, BK002 dengan BK002
            uid = no_pusat  # ID = No Pusat
            users.append({
                "id": uid,  # BA001, BK002, BH562 dll
                "nama": f"KP {no_pusat} - {nama_pusat[:40]}",
                "peranan": "KP",
                "ppd": ppd,
                "no_pusat": no_pusat,
                "nama_pusat": nama_pusat,
                "password": no_pusat.lower(),  # ba001
                "password_alt": no_pusat.upper(),  # BA001
                "password_alt2": f"kp{idx:03d}".lower(),  # legacy kp001 masih boleh
                "telefon": "",
                "kod_ppd": no_pusat[:2] if len(no_pusat)>=2 else ""
            })
            # Juga tambah alias KP001 -> BA001 untuk backward compatibility
            users.append({
                "id": f"KP{idx:03d}",
                "nama": f"KP {no_pusat} - {nama_pusat[:30]} (Alias)",
                "peranan": "KP",
                "ppd": ppd,
                "no_pusat": no_pusat,
                "nama_pusat": nama_pusat,
                "password": f"kp{idx:03d}".lower(),
                "password_alt": no_pusat.lower(),
                "password_alt2": no_pusat.upper(),
                "telefon": "",
                "kod_ppd": no_pusat[:2] if len(no_pusat)>=2 else ""
            })
    else:
        # Fallback 504 - ID = No Pusat juga
        for i in range(1, 505):
            ppd = PPD_LIST[(i-1) % 10]
            kod = list(KOD_PPD_MAPPING.keys())[(i-1) // 60 % len(KOD_PPD_MAPPING)]
            num = (i-1) % 60 + 1
            no_pusat = f"{kod}{num:03d}"
            nama_pusat = f"SMK {ppd} {num}"
            if i==1:
                no_pusat="BH562"
                nama_pusat="MAAHAD SAINS AL-QURAN AL-AMIN"
                ppd="Petaling Perdana"
            elif i==2:
                no_pusat="BC001"
                nama_pusat="SMK Sepang"
                ppd="Sepang"
            users.append({
                "id": no_pusat,
                "nama": f"KP {no_pusat} - {nama_pusat[:30]}",
                "peranan": "KP",
                "ppd": ppd,
                "no_pusat": no_pusat,
                "nama_pusat": nama_pusat,
                "password": no_pusat.lower(),
                "password_alt": no_pusat.upper(),
                "telefon": "",
                "kod_ppd": kod
            })

    # Legacy BH562 / BC001 tetap boleh login
    users.append({"id":"BH562","nama":"MAAHAD SAINS AL-QURAN AL-AMIN (Legacy)","peranan":"KP","ppd":"Petaling Perdana","no_pusat":"BH562","nama_pusat":"MAAHAD SAINS AL-QURAN AL-AMIN","password":"bh562","telefon":""})
    users.append({"id":"BC001","nama":"SMK Sepang (Legacy)","peranan":"KP","ppd":"Sepang","no_pusat":"BC001","nama_pusat":"SMK Sepang","password":"bc001","telefon":""})
    return users


def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)

# OPTIMIZE + AUTO MIGRATION - FIX Petaling Utama bug - TAK LARI dari app_15.py asal
@st.cache_data(ttl=60, show_spinner=False)
def load_reports_cached(file_mtime):
    if os.path.exists(REPORTS_FILE):
        try:
            with open(REPORTS_FILE, "r") as f:
                reports = json.load(f)
                # FIX: Jangan overwrite PPD valid - kekalkan PPD asal dari KP login
                for r in reports:
                    no_pusat = str(r.get("no_pusat","")).upper().strip()
                    kod = no_pusat[:2] if len(no_pusat)>=2 else ""
                    ppd_lama = str(r.get("ppd","")).strip()
                    # Hanya fix jika kosong atau masih kod
                    if not ppd_lama and kod in KOD_PPD_MAPPING:
                        r["ppd"] = KOD_PPD_MAPPING[kod]
                        r["kod_ppd"] = kod
                    elif ppd_lama in KOD_PPD_MAPPING:
                        r["ppd"] = KOD_PPD_MAPPING[ppd_lama]
                        r["kod_ppd"] = ppd_lama
                    if not r.get("kod_ppd") and kod in KOD_PPD_MAPPING:
                        r["kod_ppd"] = kod
                return reports
        except:
            return []
    return []

def load_reports():
    # V60 FINAL: Try Google Sheets first - data kekal walau reboot
    try:
        gs_data = load_reports_from_gsheets()
        if gs_data is not None and len(gs_data) > 0:
            return gs_data
    except:
        pass
    # Fallback JSON local
    try:
        mtime = os.path.getmtime(REPORTS_FILE) if os.path.exists(REPORTS_FILE) else 0
    except:
        mtime = 0
    return load_reports_cached(mtime)

def save_reports(reports):
    with open(REPORTS_FILE, "w") as f:
        json.dump(reports, f, indent=2)
    try:
        load_reports_cached.clear()
    except:
        pass

def clear_reports_cache():
    try:
        load_reports_cached.clear()
    except:
        pass

# ========== FUNGSI JANA BORANG BLPP HTML (OFFICIAL) ==========





def jana_blpp_html(report):
    no_pusat = report.get('no_pusat','')
    nama_pusat = report.get('nama_pusat','')
    nama_kp = report.get('nama_kp','')
    peperiksaan = report.get('peperiksaan','').upper()
    tahun = report.get('tahun','')
    kes_data = report.get('kes_data', [])
    tiada_kes = report.get('tiada_kes', False)
    pelapor_html = report.get('pelapor','') or ""
    logo_b64 = ""
    try:
        import base64, os
        for lp in ["logo_lp.png", "data/logo_lp.png", "./logo_lp.png", "/mnt/data/logo_lp.png"]:
            if os.path.exists(lp):
                try:
                    with open(lp, "rb") as f:
                        logo_b64 = base64.b64encode(f.read()).decode()
                        break
                except:
                    pass
        if not logo_b64 and os.path.exists("/tmp/logo_final.png"):
            try:
                with open("/tmp/logo_final.png", "rb") as f:
                    logo_b64 = base64.b64encode(f.read()).decode()
            except:
                pass
        if not logo_b64:
            logo_b64 = "[STRIPPED 92 bytes]"
    except:
        logo_b64 = "[STRIPPED 92 bytes]"

    
    sig_b64_html = report.get('tandatangan_b64','')
    if sig_b64_html:
        sig_img_html = f'<img src="data:image/png;base64,{sig_b64_html}" style="max-height:55px; max-width:200px;">'
    else:
        sig_img_html = '<div style="height:55px;"></div>'
    
    tarikh_html_raw = report.get('tarikh','')
    try:
        from datetime import datetime as dth
        tarikh_html_obj = dth.strptime(tarikh_html_raw, "%Y-%m-%d")
        tarikh_html = tarikh_html_obj.strftime("%d/%m/%Y")
    except:
        tarikh_html = tarikh_html_raw if tarikh_html_raw else "___/___/______"
    
    if peperiksaan and tahun:
        pep_line = f"PEPERIKSAAN {peperiksaan} TAHUN {tahun}"
    else:
        pep_line = "PEPERIKSAAN ........................................ TAHUN ........................................"
    
    if tiada_kes or not kes_data:
        kes_rows = "<tr><td style='height:32px; text-align:center;'>1</td><td colspan='4' style='text-align:center; font-weight:bold; font-size:11pt; height:32px;'>TIADA KES LUAR DUGAAN - TIADA KES PADA HARI INI</td></tr>"
        for _ in range(1):
            kes_rows += "<tr><td style='height:20px;'>&nbsp;</td><td></td><td></td><td></td><td></td></tr>"
    else:
        kes_rows = ""
        for i, k in enumerate(kes_data, 1):
            kes_rows += f"<tr><td style='height:auto; min-height:32px; word-wrap:break-word; white-space:pre-wrap;'>{i}</td><td style='word-wrap:break-word; white-space:pre-wrap;'>{k.get('angka_giliran','')}</td><td style='word-wrap:break-word; white-space:pre-wrap;'>{k.get('no_ic','')}</td><td style='text-align:left; word-wrap:break-word; white-space:pre-wrap; word-break:break-word;'>{k.get('kes','')}</td><td style='text-align:left; word-wrap:break-word; white-space:pre-wrap; word-break:break-word;'>{k.get('tindakan','')}</td></tr>"
        for _ in range(max(0, 5 - len(kes_data))):
            kes_rows += "<tr><td style='min-height:32px; height:auto;'>&nbsp;</td><td></td><td></td><td></td><td></td></tr>"
    
    html = f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
        @page {{ size: A4 landscape; margin: 10mm 12mm 8mm 12mm; }}
        body {{ font-family: Arial, sans-serif; margin:0; color:#000; }}
        .page {{ width:100%; max-width:1020px; margin:0 auto; }}
        .hdr {{ display:flex; justify-content:space-between; align-items:flex-start; }}
        .hdr-center {{ text-align:center; font-weight:bold; font-size:11pt; flex:1; line-height:1.25; }}
        .hdr-right {{ text-align:right; font-size:9pt; font-weight:bold; line-height:1.2; }}
        hr {{ border:none; border-top:1.8px solid #000; margin:8px 0; }}
        .pep {{ text-align:center; font-weight:bold; font-size:11pt; margin:8px 0 2px 0; }}
        .title {{ text-align:center; font-weight:bold; font-size:13pt; margin:2px 0 10px 0; }}
        .arahan {{ width:100%; font-size:8.2pt; text-align:justify; line-height:1.35; box-sizing:border-box; margin:0; padding:0; }}
        .arahan ol {{ margin:0; padding-left:14px; width:100%; text-align:justify; }}
        .arahan li {{ text-align:justify; width:100%; }}
        .arahan ol {{ margin:0; padding-left:18px; }}
        table {{ width:100%; border-collapse:collapse; table-layout:fixed; margin:0; border-spacing:0; }}
        table + table {{ margin-top:0 !important; }}
        th, td {{ border:1.2px solid #000; }}
        th, td {{ word-wrap:break-word; white-space:pre-wrap; word-break:break-word; border:1.2px solid #000; padding:4px 5px; font-size:9pt; }}
        th {{ text-align:center; font-weight:bold; vertical-align:middle; }}
        td {{ text-align:center; vertical-align:middle; }}
        .footer {{ display:flex; justify-content:space-between; margin-top:26px; font-size:9pt; }}
    </style></head><body>
    <div class="page">
        <div class="hdr">
            <div style="width:100px;"><img src="data:image/png;base64,{logo_b64}" style="width:85px; height:auto; display:block;" onerror="this.style.display='none'"></div>
            <div class="hdr-center">LEMBAGA PEPERIKSAAN<br>KEMENTERIAN PENDIDIKAN MALAYSIA</div>
            <div class="hdr-right">LP/Am 90 Pin. 1/2023<br><b style="font-size:13pt;">BLPP</b></div>
        </div>
        <!-- hr dibuang ikut request biru -->
        <div class="pep">{pep_line}</div>
        <div class="title">BORANG LAPORAN PENGENDALIAN PEPERIKSAAN</div>
        <div class="arahan"><ol>
            <li>Borang ini hendaklah digunakan semasa peperiksaan untuk melaporkan mengenai perkara-perkara di luar dugaan yang berlaku kepada calon seperti sakit secara tiba-tiba, kemalangan, ditahan polis dan lain-lain.</li>
            <li>Jika ada kes, Ketua Pengawas Peperiksaan/Ketua Pengawas Amali/Ketua Pentaksir/Assessor dikehendaki mengisi borang ini dan menyerahkannya kepada Penyelia Kawasan untuk dihantar melalui e-mel pada hari yang sama kepada Pengarah Peperiksaan, Lembaga Peperiksaan (u.p.: Unit Menengah Atas/Unit Pra Universiti/Unit Vokasional) dan Pengarah Pendidikan Negeri, Jabatan Pendidikan Negeri (u.p.: Sektor Pentaksiran dan Peperiksaan).</li>
        </ol></div>
        <table style="margin-top:10px; margin-bottom:0px; border-bottom:none; border-collapse:collapse;">
            <colgroup>
                <col style="width:6%;">
                <col style="width:16%;">
                <col style="width:16%;">
                <col style="width:31%;">
                <col style="width:31%;">
            </colgroup>
            <tr style="height:32px;">
                <th colspan="2" style="width:22%; font-size:8pt; padding:2px 3px; height:32px; line-height:1.15;">NO. PUSAT</th>
                <th colspan="2" style="width:47%; font-size:8pt; padding:2px 3px; height:32px; line-height:1.15;">NAMA PUSAT PEPERIKSAAN</th>
                <th style="width:31%; font-size:6.8pt; line-height:1.15; text-align:center; vertical-align:middle; padding:2px 3px; height:32px; font-weight:bold;">
                    <div style="text-align:center; margin:0; padding:0; line-height:1.15;">NAMA KETUA PENGAWAS</div>
                    <div style="text-align:center; margin:0; padding:0; line-height:1.15;">PEPERIKSAAN / KETUA PENGAWAS</div>
                    <div style="text-align:center; margin:0; padding:0; line-height:1.15;">AMALI / KETUA PENTAKSIR /</div>
                    <div style="text-align:center; margin:0; padding:0; line-height:1.15;">ASSESSOR</div>
                </th>
            </tr>
            <tr style="height:22px;">
                <td colspan="2" style="font-weight:bold; height:22px; border-bottom:none; text-align:center; font-size:8.5pt; padding:2px;">{no_pusat}</td>
                <td colspan="2" style="font-weight:bold; border-bottom:none; text-align:center; font-size:8.5pt; padding:2px;">{nama_pusat}</td>
                <td style="font-weight:bold; word-wrap:break-word; border-bottom:none; text-align:center; font-size:8.5pt; padding:2px; height:22px;">{nama_kp}</td>
            </tr>
        </table>
        <table style="margin-top:0px; border-top:none; border-collapse:collapse;">
            <colgroup>
                <col style="width:6%;">
                <col style="width:16%;">
                <col style="width:16%;">
                <col style="width:31%;">
                <col style="width:31%;">
            </colgroup>
            <tr>
                <th style="width:6%;">BIL.</th>
                <th style="width:16%;">ANGKA GILIRAN</th>
                <th style="width:16%;">NO. PENGENALAN DIRI</th>
                <th style="width:31%;">KES</th>
                <th style="width:31%;">TINDAKAN</th>
            </tr>
            {kes_rows}
        </table>
                <div class="footer" style="display:flex; justify-content:space-between;">
            <div style="width:45%;">Pelapor:<br><div style="height:50px; margin:5px 0;">{sig_img_html}</div>
                <div style="border-top:1px solid black; margin-top:5px; padding-top:5px;">
                    <div style="font-weight:800; font-size:14px; text-transform:uppercase;">{pelapor_html}</div>
                    <div style="font-size:10px; margin-top:4px;">Tandatangan Ketua Pengawas Peperiksaan/<br>Ketua Pengawas Amali/Ketua Pentaksir/Assessor</div>
                    <div style="font-size:11px; margin-top:8px;">Tarikh: {tarikh_html}</div>
                </div>
            </div>
            <div style="width:45%;">Pengesahan:<br><br><br>
                <div style="border-top:1px solid black; margin-top:50px; padding-top:5px;">
                    <div style="font-size:10px;">Tandatangan Penyelia Kawasan/Pegawai Peperiksaan</div>
                    <div style="font-size:11px; margin-top:8px;">Tarikh: {tarikh_html}</div>
                </div>
            </div>
        </div>
    </div></body></html>
    """
    return html
def create_bulk_pdf_zip_by_date(reports, tarikh_selected):
    import zipfile
    from io import BytesIO
    filtered = [r for r in reports if r.get('tarikh') == tarikh_selected]
    if not filtered:
        return None, 0
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for r in filtered:
            no_pusat = r.get('no_pusat','UNKNOWN')
            safe_nama = "".join(c if c.isalnum() or c in ('-','_') else '_' for c in r.get('nama_pusat','')[:15])
            base_filename = f"{r.get('tarikh','')}_{no_pusat}_{safe_nama}"
            try:
                pdf_data = jana_blpp_pdf(r)
                if isinstance(pdf_data, BytesIO):
                    pdf_bytes = pdf_data.getvalue()
                else:
                    pdf_bytes = pdf_data
                zip_file.writestr(f"{base_filename}.pdf", pdf_bytes)
            except:
                try:
                    html_content = jana_blpp_html(r)
                    zip_file.writestr(f"{base_filename}.html", html_content)
                except:
                    pass
    zip_buffer.seek(0)
    return zip_buffer.getvalue(), len(filtered)



def jana_blpp_pdf(report):
    """PDF 100% IKUT TEMPLATE ASAL LP/Am 90 Pin. 1/2023 - BLPP"""
    no_pusat = report.get('no_pusat','')
    nama_pusat = report.get('nama_pusat','')
    nama_kp = report.get('nama_kp','')
    pep = report.get('peperiksaan','')
    thn = report.get('tahun','')
    kes_data = report.get('kes_data', [])
    tiada = report.get('tiada_kes', False)

    try:
        from fpdf import FPDF
        from PIL import Image
        import os

        logo_to_use = None
        for lp in ["logo_lp.png", "data/logo_lp.png", "./logo_lp.png", "/tmp/logo_final.png", "LEMBAGA PEPERIKSAAN.png"]:
            if os.path.exists(lp):
                try:
                    im = Image.open(lp)
                    tmp = "/tmp/logo_final.png"
                    im.convert("RGBA").save(tmp, "PNG")
                    logo_to_use = tmp
                    break
                except:
                    logo_to_use = lp
                    break

        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=False)
        pdf.add_page()
        pdf.set_line_width(0.35)

        if logo_to_use:
            try:
                pdf.image(logo_to_use, x=10, y=4, w=28)
            except:
                pass

        # Header tengah
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_xy(42, 6)
        pdf.cell(190, 5, "LEMBAGA PEPERIKSAAN", align="C")
        pdf.set_xy(42, 11)
        pdf.cell(190, 5, "KEMENTERIAN PENDIDIKAN MALAYSIA", align="C")

        # Kanan atas
        pdf.set_font("Helvetica", "", 8)
        pdf.set_xy(250, 6)
        pdf.cell(37, 4, "LP/Am 90 Pin. 1/2023", align="R")
        pdf.set_font("Helvetica", "B", 14)
        pdf.set_xy(250, 11)
        pdf.cell(37, 7, "BLPP", align="R")

        # Garis atas dibuang ikut request - tiada line biru

        pep_text = f"PEPERIKSAAN {pep} TAHUN {thn}" if pep and thn else "PEPERIKSAAN ........................................ TAHUN ........................................"
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_xy(10, 22)
        pdf.cell(277, 5, pep_text, align="C")
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_xy(10, 28)
        pdf.cell(277, 6, "BORANG LAPORAN PENGENDALIAN PEPERIKSAAN", align="C")

        pdf.set_font("Helvetica", "", 8.2)
        pdf.set_xy(10, 36)
        pdf.set_font("Helvetica", "", 7.2)
        arahan_full_1 = "1.  Borang ini hendaklah digunakan semasa peperiksaan untuk melaporkan mengenai perkara-perkara di luar dugaan yang berlaku kepada calon seperti sakit secara tiba-tiba, kemalangan, ditahan polis dan lain-lain."
        arahan_full_2 = "2.  Jika ada kes, Ketua Pengawas Peperiksaan/Ketua Pengawas Amali/Ketua Pentaksir/Assessor dikehendaki mengisi borang ini dan menyerahkannya kepada Penyelia Kawasan untuk dihantar melalui e-mel pada hari yang sama kepada Pengarah Peperiksaan, Lembaga Peperiksaan (u.p.: Unit Menengah Atas/Unit Pra Universiti/Unit Vokasional) dan Pengarah Pendidikan Negeri, Jabatan Pendidikan Negeri (u.p.: Sektor Pentaksiran dan Peperiksaan)."
        pdf.set_xy(10, 36)
        pdf.multi_cell(277, 3.2, arahan_full_1, align="J")
        pdf.set_xy(10, 43)
        pdf.multi_cell(277, 3.2, arahan_full_2, align="J")

        # ===== TABLE 1 - IKUT TEMPLATE ASAL =====
        # Total 277mm : 20% = 55.4, 38% = 105.26, 42% = 116.34
        y1 = 56
        w1, w2, w3 = 60, 131, 86  # FIX: Garis menegak lurus
        h_header = 14

        pdf.set_font("Helvetica", "B", 8.5)
        # Header row
        pdf.set_xy(10, y1)
        pdf.cell(w1, h_header, "NO. PUSAT", border=1, align="C")
        pdf.cell(w2, h_header, "NAMA PUSAT PEPERIKSAAN", border=1, align="C")
        # Kotak ketiga dengan 3 baris centre
        x3 = 10 + w1 + w2
        pdf.set_xy(x3, y1)
        pdf.cell(w3, h_header, "", border=1)
        pdf.set_font("Helvetica", "B", 7.2)
        pdf.set_xy(x3, y1+0.8)
        pdf.cell(w3, 4, "NAMA KETUA PENGAWAS", align="C")
        pdf.set_xy(x3, y1+4.8)
        pdf.cell(w3, 4, "PEPERIKSAAN/KETUA PENGAWAS", align="C")
        pdf.set_xy(x3, y1+8.8)
        pdf.cell(w3, 4, "AMALI/KETUA PENTAKSIR/ASSESSOR", align="C")

        # Data row
        h_data = 10
        pdf.set_xy(10, y1+h_header)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(w1, h_data, no_pusat, border=1, align="C")
        pdf.set_font("Helvetica", "B", 9)
        # Potong jika panjang tapi centre
        pdf.cell(w2, h_data, nama_pusat[:50], border=1, align="C")
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(w3, h_data, nama_kp[:65], border=1, align="C")

        # ===== TABLE 2 - FIX GARISAN MERAH SAMA ATAS BAWAH =====
        # Pastikan outer width sama 277mm macam table atas: 10 ke 287
        # Dan semua vertical line guna x yang exact, tak ada rounding error
        y2 = y1 + h_header + h_data + 0  # FIX: Tiada jarak, rapat
        # Guna width yang buat vertical line bawah lurus sama dari header sampai bawah
        wb1, wb2, wb3, wb4, wb5 = 16, 44, 45, 86, 86  # total 277
        h2_header = 10
        # Define x positions sekali, guna untuk header dan body - baru sama
        x0 = 10
        x1 = x0 + wb1
        x2 = x1 + wb2
        x3_b = x2 + wb3
        x4 = x3_b + wb4
        x5 = x4 + wb5  # should be 287

        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_line_width(0.35)  # pastikan line width sama semua
        # Header - draw dengan x exact
        pdf.set_xy(x0, y2)
        pdf.cell(wb1, h2_header, "BIL.", border=1, align="C")
        pdf.set_xy(x1, y2)
        pdf.cell(wb2, h2_header, "ANGKA GILIRAN", border=1, align="C")
        # NO PENGENALAN DIRI 2 baris - dalam kotak yang sama border
        pdf.set_xy(x2, y2)
        pdf.cell(wb3, h2_header, "", border=1)
        pdf.set_font("Helvetica", "B", 7)
        pdf.set_xy(x2, y2+1)
        pdf.cell(wb3, 4, "NO. PENGENALAN", align="C")
        pdf.set_xy(x2, y2+5)
        pdf.cell(wb3, 4, "DIRI", align="C")
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_xy(x3_b, y2)
        pdf.cell(wb4, h2_header, "KES", border=1, align="C")
        pdf.set_xy(x4, y2)
        pdf.cell(wb5, h2_header, "TINDAKAN", border=1, align="C")

        # ===== FLEXIBLE BOX - Kotak auto panjang ikut tulisan (FIX terkeluar kotak) =====
        y_body = y2 + h2_header
        
        def get_wrapped_lines(pdf_obj, text, col_width):
            # Kira berapa baris sebenar bila wrap ikut lebar column - accurate
            if not text:
                return [""]
            lines = []
            # Split by manual newlines first
            for para in str(text).split("\n"):
                para = para.strip()
                if not para:
                    lines.append("")
                    continue
                words = para.split()
                cur_line = ""
                for w in words:
                    test = cur_line + " " + w if cur_line else w
                    if pdf_obj.get_string_width(test) < col_width - 3:
                        cur_line = test
                    else:
                        if cur_line:
                            lines.append(cur_line)
                        # Jika satu perkataan panjang sangat, potong
                        if pdf_obj.get_string_width(w) >= col_width - 3:
                            # potong perkataan panjang
                            temp = w
                            while temp:
                                for i in range(len(temp), 0, -1):
                                    if pdf_obj.get_string_width(temp[:i]) < col_width - 3:
                                        lines.append(temp[:i])
                                        temp = temp[i:]
                                        break
                                else:
                                    lines.append(temp)
                                    break
                            cur_line = ""
                        else:
                            cur_line = w
                if cur_line:
                    lines.append(cur_line)
            return lines if lines else [""]

        def calc_text_height(pdf_obj, text, col_width, line_h=4):
            lines = get_wrapped_lines(pdf_obj, text, col_width-2)
            return max(8, len(lines) * line_h + 4)
        
        if tiada or not kes_data:
            big_h = 28
            pdf.set_xy(x0, y_body)
            pdf.cell(wb1, big_h, "1", border=1, align="C")
            pdf.set_xy(x1, y_body)
            pdf.cell(wb2, big_h, "", border=1)
            pdf.set_xy(x2, y_body)
            pdf.cell(wb3, big_h, "", border=1)
            pdf.set_font("Helvetica", "B", 11)
            pdf.set_xy(x3_b, y_body)
            pdf.cell(wb4+wb5, big_h, "TIADA KES LUAR DUGAAN - TIADA KES PADA HARI INI", border=1, align="C")
            total_table_h = big_h
        else:
            # ADA KES - Lukis kotak flexible ikut panjang tulisan
            row_y = y_body
            total_table_h = 0
            row_heights = []
            
            # First pass: kira tinggi setiap baris dengan accurate wrap
            for k in kes_data:
                kes_text = str(k.get('kes','')).strip()
                tind_text = str(k.get('tindakan','')).strip()
                h_kes = calc_text_height(pdf, kes_text, wb4-2, 4)
                h_tind = calc_text_height(pdf, tind_text, wb5-2, 4)
                row_h = max(14, h_kes, h_tind, 14)
                row_heights.append(row_h)
                total_table_h += row_h
            
            # Jangan paksa 62mm kalau 1 kes sahaja - biar flexible (request Kashah)
            # total_table_h kekal ikut panjang sebenar
            
            # Second pass: lukis setiap baris dengan border flexible + ISI PENUH PAGE 1 DAHULU
            row_y = y_body
            page_limit = 195  # Guna penuh page 1 sampai hampir bawah (A4 landscape 210mm - margin)
            for idx, k in enumerate(kes_data):
                kes_text = str(k.get('kes','')).strip()
                tind_text = str(k.get('tindakan','')).strip()
                pre_row_h = row_heights[idx]
                
                # PENUHKAN PAGE 1 DAHULU: Kira baki ruang di page 1
                remaining_space = page_limit - row_y
                
                # Dapatkan lines untuk KES dan TINDAKAN
                kes_lines_all = get_wrapped_lines(pdf, str(k.get('kes','')).strip(), wb4-2)
                tind_lines_all = get_wrapped_lines(pdf, str(k.get('tindakan','')).strip(), wb5-2)
                total_lines_needed = max(len(kes_lines_all), len(tind_lines_all))
                pre_row_h = max(14, total_lines_needed * 4 + 6)
                
                # Jika row muat di page semasa, lukis terus
                # Jika tidak muat TAPI masih ada ruang > 25mm di page 1, kita split - isi penuh page 1 dulu
                need_split = pre_row_h > remaining_space and remaining_space > 25 and row_y > y_body  # ada isi di page 1
                
                if pre_row_h > page_limit - row_y and row_y + 15 > page_limit - 30:
                    # Memang tak muat langsung, buat page baru terus
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 8)
                    pdf.set_xy(10, 10)
                    pdf.cell(277, 5, f"SAMBUNGAN - {no_pusat} | Page {pdf.page_no()}", align="C")
                    row_y = 18
                    pdf.set_font("Helvetica", "B", 7.5)
                    pdf.set_xy(x0, row_y)
                    pdf.cell(wb1, 8, "BIL.", border=1, align="C")
                    pdf.set_xy(x1, row_y)
                    pdf.cell(wb2, 8, "ANGKA GILIRAN", border=1, align="C")
                    pdf.set_xy(x2, row_y)
                    pdf.cell(wb3, 8, "NO. PENGENALAN DIRI", border=1, align="C")
                    pdf.set_xy(x3_b, row_y)
                    pdf.cell(wb4, 8, "KES (Sambungan)", border=1, align="C")
                    pdf.set_xy(x4, row_y)
                    pdf.cell(wb5, 8, "TINDAKAN (Sambungan)", border=1, align="C")
                    row_y += 8
                    # Recalc remaining after new page
                    remaining_space = page_limit - row_y
                    need_split = False
                
                pdf.set_line_width(0.35)
                start_y = row_y
                
                # Jika perlu split, kita akan lukis sebahagian di page 1, baki di page 2
                if need_split:
                    # Kira berapa baris muat di page 1
                    lines_fit = max(1, int((remaining_space - 6) / 4))
                    kes_part1 = "\n".join(kes_lines_all[:lines_fit])
                    kes_part2 = "\n".join(kes_lines_all[lines_fit:])
                    tind_part1 = "\n".join(tind_lines_all[:lines_fit])
                    tind_part2 = "\n".join(tind_lines_all[lines_fit:])
                    
                    # Lukis part 1 di page 1 - PENUHKAN PAGE 1
                    part1_h = remaining_space
                    pdf.rect(x0, start_y, wb1, part1_h)
                    pdf.rect(x1, start_y, wb2, part1_h)
                    pdf.rect(x2, start_y, wb3, part1_h)
                    pdf.rect(x3_b, start_y, wb4, part1_h)
                    pdf.rect(x4, start_y, wb5, part1_h)
                    
                    pdf.set_font("Helvetica", "", 7)
                    pdf.set_xy(x0, start_y+1)
                    pdf.cell(wb1, 5, str(idx+1), align="C")
                    pdf.set_xy(x1, start_y+1)
                    pdf.cell(wb2, 5, str(k.get('angka_giliran',''))[:20], align="C")
                    pdf.set_xy(x2, start_y+1)
                    pdf.cell(wb3, 5, str(k.get('no_ic',''))[:20], align="C")
                    pdf.set_xy(x3_b+1, start_y+1)
                    pdf.multi_cell(wb4-2, 4, kes_part1, align="L")
                    pdf.set_xy(x4+1, start_y+1)
                    pdf.multi_cell(wb5-2, 4, tind_part1, align="L")
                    
                    # Page baru untuk sambungan
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 8)
                    pdf.set_xy(10, 10)
                    pdf.cell(277, 5, f"SAMBUNGAN - {no_pusat} | Page {pdf.page_no()} - PENUHKAN PAGE 1 SELESAI", align="C")
                    row_y = 18
                    pdf.set_font("Helvetica", "B", 7.5)
                    pdf.set_xy(x0, row_y)
                    pdf.cell(wb1, 8, "BIL.", border=1, align="C")
                    pdf.set_xy(x1, row_y)
                    pdf.cell(wb2, 8, "ANGKA GILIRAN", border=1, align="C")
                    pdf.set_xy(x2, row_y)
                    pdf.cell(wb3, 8, "NO. PENGENALAN DIRI", border=1, align="C")
                    pdf.set_xy(x3_b, row_y)
                    pdf.cell(wb4, 8, "KES (Sambungan)", border=1, align="C")
                    pdf.set_xy(x4, row_y)
                    pdf.cell(wb5, 8, "TINDAKAN (Sambungan)", border=1, align="C")
                    row_y += 8
                    start_y = row_y
                    
                    # Kira tinggi baki
                    kes_lines_remain = get_wrapped_lines(pdf, kes_part2, wb4-2)
                    tind_lines_remain = get_wrapped_lines(pdf, tind_part2, wb5-2)
                    remain_h = max(14, max(len(kes_lines_remain), len(tind_lines_remain)) * 4 + 6)
                    
                    # Lukis baki di page 2
                    pdf.rect(x0, start_y, wb1, remain_h)
                    pdf.rect(x1, start_y, wb2, remain_h)
                    pdf.rect(x2, start_y, wb3, remain_h)
                    pdf.rect(x3_b, start_y, wb4, remain_h)
                    pdf.rect(x4, start_y, wb5, remain_h)
                    
                    pdf.set_font("Helvetica", "", 7)
                    pdf.set_xy(x0, start_y+1)
                    pdf.cell(wb1, 5, f"{idx+1} samb.", align="C")
                    pdf.set_xy(x3_b+1, start_y+1)
                    pdf.multi_cell(wb4-2, 4, kes_part2, align="L")
                    pdf.set_xy(x4+1, start_y+1)
                    pdf.multi_cell(wb5-2, 4, tind_part2, align="L")
                    
                    row_y = start_y + remain_h
                    continue  # dah siap split, next kes
                
                
                # Isi data - untuk kira actual height
                pdf.set_font("Helvetica", "", 7)
                
                # BIL, AG, IC - tulis terus
                pdf.set_xy(x0, start_y+1)
                pdf.cell(wb1, 5, str(idx+1), align="C")
                pdf.set_xy(x1, start_y+1)
                pdf.cell(wb2, 5, str(k.get('angka_giliran',''))[:20], align="C")
                pdf.set_xy(x2, start_y+1)
                pdf.cell(wb3, 5, str(k.get('no_ic',''))[:20], align="C")
                
                # KES
                pdf.set_xy(x3_b+1, start_y+1)
                y_before_kes = pdf.get_y()
                pdf.multi_cell(wb4-2, 4, kes_text, align="L")
                y_after_kes = pdf.get_y()
                
                # TINDAKAN
                pdf.set_xy(x4+1, start_y+1)
                y_before_tind = pdf.get_y()
                pdf.multi_cell(wb5-2, 4, tind_text, align="L")
                y_after_tind = pdf.get_y()
                
                # Actual height used
                actual_h = max(pre_row_h, y_after_kes - start_y, y_after_tind - start_y, 14)
                actual_h = max(actual_h, 14) + 2  # padding
                
                # Sekarang lukis border dengan actual height (overwrite)
                pdf.rect(x0, start_y, wb1, actual_h)
                pdf.rect(x1, start_y, wb2, actual_h)
                pdf.rect(x2, start_y, wb3, actual_h)
                pdf.rect(x3_b, start_y, wb4, actual_h)
                pdf.rect(x4, start_y, wb5, actual_h)
                
                row_y = start_y + actual_h
            
            # FIX: Kalau 1 kes sahaja, tak perlu tambah 5 kotak kosong - biar 1 baris sahaja (request Kashah)
            # Hanya tambah kosong jika tiada kes langsung
            # Kalau ada 1 kes, terus guna tinggi sebenar sahaja - tak perlu isi kosong
            # total_table_h sudah dikira dari row_heights, kekalkan

        page_limit = 185
        try:
            curr_bottom = row_y
        except:
            try:
                curr_bottom = y_body + total_table_h
            except:
                try:
                    curr_bottom = y_body + big_h
                except:
                    curr_bottom = y_body + 28
        
        if curr_bottom + 35 > page_limit:
            pdf.add_page()
            # Header ringkas untuk page 2 (sambungan)
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_xy(10, 10)
            pdf.cell(277, 5, f"SAMBUNGAN LAPORAN - {no_pusat} - {nama_pusat} - {len(kes_data)} KES", align="C")
            yf = 18
        else:
            yf = curr_bottom + 8
        # ===== FOOTER BARU - NAMA KETUA PENGAWAS TURUN & BESAR - FIX KASHAH =====
        pelapor_name = report.get('pelapor','') or ""  # Kosong jika KP tak isi - jangan guna nama sekolah
        
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_xy(10, yf)
        pdf.cell(60, 5, "Pelapor:")
        pdf.set_xy(150, yf)
        pdf.cell(60, 5, "Pengesahan:")
        
        # Signature clear background - letak di atas garisan, tidak tutup text
        sig_b64 = report.get('tandatangan_b64','')
        if sig_b64:
            try:
                import base64
                img_bytes = base64.b64decode(sig_b64)
                import os, tempfile
                tmp_path = os.path.join(tempfile.gettempdir(), f"sig_{report.get('id','tmp')}.png")
                with open(tmp_path, 'wb') as f:
                    f.write(img_bytes)
                # Letak signature di ruang kosong antara Pelapor: dan garisan
                pdf.image(tmp_path, x=14, y=yf+5, w=38, h=11)
                try:
                    os.remove(tmp_path)
                except:
                    pass
            except Exception as e:
                pass
        
        # Garisan tandatangan
        pdf.line(10, yf+18, 95, yf+18)
        pdf.line(150, yf+18, 235, yf+18)
        
        # NAMA KETUA PENGAWAS - BESAR & TURUN BAWAH GARISAN (tidak tutup text belakang)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_xy(10, yf+20)
        pdf.cell(85, 6, pelapor_name.upper() if pelapor_name else "", align="L")
        
        # Label jawatan - kecil di bawah nama, tidak bertindih
        pdf.set_font("Helvetica", "", 7)
        pdf.set_xy(10, yf+26)
        pdf.cell(85, 4, "Tandatangan Ketua Pengawas Peperiksaan/")
        pdf.set_xy(10, yf+30)
        pdf.cell(85, 4, "Ketua Pengawas Amali/Ketua Pentaksir/Assessor")
        
        pdf.set_xy(150, yf+26)
        pdf.cell(85, 4, "Tandatangan Penyelia Kawasan/")
        pdf.set_xy(150, yf+30)
        pdf.cell(85, 4, "Pegawai Peperiksaan")
        
        # Tarikh auto
        tarikh_laporan = report.get('tarikh','')
        try:
            from datetime import datetime as dt
            tarikh_obj_pdf = dt.strptime(tarikh_laporan, "%Y-%m-%d")
            tarikh_pdf_formatted = tarikh_obj_pdf.strftime("%d/%m/%Y")
        except:
            tarikh_pdf_formatted = tarikh_laporan if tarikh_laporan else "___/___/______"
        
        pdf.set_font("Helvetica", "", 8)
        pdf.set_xy(10, yf+36)
        pdf.cell(80, 5, f"Tarikh: {tarikh_pdf_formatted}")
        pdf.set_xy(150, yf+36)
        pdf.cell(80, 5, f"Tarikh: {tarikh_pdf_formatted}")

        out = pdf.output(dest='S')
        return out.encode('latin1') if isinstance(out, str) else bytes(out)
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            from fpdf import FPDF
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("Helvetica","B",14)
            pdf.cell(0,10,f"BLPP - {no_pusat}", align="C", new_x="LMARGIN", new_y="NEXT")
            out = pdf.output(dest='S')
            return out.encode('latin1') if isinstance(out, str) else bytes(out)
        except:
            return f"BLPP {no_pusat}".encode()


# ========== LOGIN ==========
# ========== LOGIN ==========
# ========== LOGIN ==========

if "login" not in st.session_state:
    st.session_state.login = False
    st.session_state.user = None

if not st.session_state.login:
    # FRONT PAGE MENARIK - KASHAH REQUEST - BUANG DEMO LOGIN, TULIS PENUH
    st.markdown('''
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    .stApp { background: #f1f5f9 !important; }
    .main .block-container { background: #f1f5f9; }
    .header-hero {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #2563eb 100%) !important;
        border: 2px solid #1e40af;
        border-radius: 20px;
        padding: 32px 24px;
        text-align: center;
        margin-bottom: 24px;
        box-shadow: 0 10px 40px rgba(30,58,138,0.4);
    }
    .header-hero h1 {
        color: #ffffff !important;
        text-shadow: 0 2px 12px rgba(0,0,0,0.5) !important;
        font-weight: 900 !important;
    }
    .header-hero p {
        color: #dbeafe !important;
    }
    .login-card {
        background: #ffffff !important;
        border-radius: 20px;
        padding: 28px;
        box-shadow: 0 12px 40px rgba(0,0,0,0.15), 0 0 0 1px #e2e8f0;
        border: 2px solid #e2e8f0;
    }
    .role-badge {
        display: inline-block;
        padding: 8px 14px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 0.3px;
        margin: 4px;
        border: 1px solid;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }
    .badge-kp { background: #eff6ff !important; color: #1d4ed8 !important; border-color: #3b82f6 !important; }
    .badge-ppd { background: #f0fdf4 !important; color: #15803d !important; border-color: #22c55e !important; }
    .badge-jpn { background: #fefce8 !important; color: #a16207 !important; border-color: #eab308 !important; }
    .badge-lp { background: #fdf2f8 !important; color: #be185d !important; border-color: #ec4899 !important; }
    .feature-box {
        background: #ffffff;
        border: 2px solid #e2e8f0;
        border-radius: 14px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    .feature-box div { color: #0f172a !important; }
    div[data-testid="stSelectbox"] > div > div { 
        border-radius: 12px !important; 
        border: 2px solid #334155 !important;
        font-weight: 700 !important;
        background: white !important;
        color: #0f172a !important;
    }
    div[data-testid="stTextInput"] > div > div > input {
        border-radius: 12px !important;
        border: 2px solid #334155 !important;
        padding: 14px !important;
        background: white !important;
        color: #0f172a !important;
        font-weight: 600 !important;
    }
    div[data-testid="stTextInput"] label, div[data-testid="stSelectbox"] label {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 14px !important;
    }
    .stButton > button {
        border-radius: 12px !important;
        padding: 14px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(220,38,38,0.4) !important;
        letter-spacing: 0.5px;
        color: white !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #b91c1c 0%, #991b1b 100%) !important;
        box-shadow: 0 6px 20px rgba(220,38,38,0.5) !important;
        transform: translateY(-1px);
    }
    </style>
    ''', unsafe_allow_html=True)
    
    # Hero Header Menarik - HIGH CONTRAST FIX KASHAH
    st.markdown('''
    <div class="header-hero">
        <div style="font-size:52px; margin-bottom:12px; filter: drop-shadow(0 4px 8px rgba(0,0,0,0.3));">📋</div>
        <h1 style="font-size:30px; font-weight:900; color:#ffffff; margin:0; letter-spacing:0.8px; font-family:Inter; text-shadow:0 3px 12px rgba(0,0,0,0.6); line-height:1.2;">
            SISTEM LAPORAN PUSAT PEPERIKSAAN
        </h1>
        <p style="font-size:16px; color:#dbeafe; margin:12px 0 0 0; font-weight:600; text-shadow:0 1px 4px rgba(0,0,0,0.4);">
            Sektor Pentaksiran dan Peperiksaan<br>
            <span style="color:#ffffff; font-weight:700;">Jabatan Pendidikan Negeri Selangor</span>
        </p>
        <div style="margin-top:20px; display:flex; justify-content:center; flex-wrap:wrap; gap:8px;">
            <span class="role-badge badge-kp">KP • KETUA PENGAWAS</span>
            <span class="role-badge badge-ppd">PPD • PEJABAT PENDIDIKAN DAERAH</span>
            <span class="role-badge badge-jpn">JPN • JABATAN PENDIDIKAN NEGERI</span>
            <span class="role-badge badge-lp">LP • LEMBAGA PEPERIKSAAN</span>
        </div>
        <div style="margin-top:18px; font-size:13px; color:#fde047; font-weight:800; letter-spacing:0.8px; background:rgba(0,0,0,0.25); padding:8px 16px; border-radius:20px; display:inline-block; border:1px solid rgba(253,224,71,0.3);">
            ✨ BLPP LP/Am 90 Pin. 1/2023 | SPM • UJIAN AMALI SAINS ✨
        </div>
    </div>
    ''', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1,1.15,1])
    with col2:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        st.markdown('<h3 style="text-align:center; margin:0 0 6px 0; font-size:20px; font-weight:800; color:#0f172a;">🔐 Log Masuk Sistem</h3><p style="text-align:center; font-size:12px; color:#64748b; margin:0 0 18px 0;">Sila pilih peranan dan masukkan ID anda</p>', unsafe_allow_html=True)
        
        users = load_users()
        # FIX KASHAH: Tulis penuh - KP KETUA PENGAWAS, PPD PEJABAT PENDIDIKAN DAERAH, JPN JABATAN PENDIDIKAN NEGERI, LP LEMBAGA PEPERIKSAAN
        peranan_options = [
            "KP - KETUA PENGAWAS",
            "PPD - PEJABAT PENDIDIKAN DAERAH",
            "JPN - JABATAN PENDIDIKAN NEGERI",
            "LP - LEMBAGA PEPERIKSAAN"
        ]
        peranan = st.selectbox("Peranan", peranan_options, label_visibility="visible")
        
        id_pengguna = st.text_input("ID Pengguna", placeholder="Contoh: KP001 / BA / JPN / LP")
        password = st.text_input("Kata Laluan", type="password", placeholder="Masukkan kata laluan")
        
        st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
        
        if st.button("Log Masuk →", type="primary", use_container_width=True):
            found = None
            uid_input = id_pengguna.strip().upper()
            pwd_input = password.strip()
            pwd_lower = pwd_input.lower()
            
            # Mapping peranan pilih ke peranan sistem
            if peranan.startswith("KP"):
                target_roles = ["KP"]
            elif peranan.startswith("PPD"):
                target_roles = ["PPD"]
            elif peranan.startswith("JPN"):
                target_roles = ["ADMIN_JPN"]
            elif peranan.startswith("LP"):
                target_roles = ["LP"]
            else:
                target_roles = []
            
            for u in users:
                # Check ID
                if u["id"].upper() != uid_input:
                    continue
                # Check password - support multiple password fields
                stored_pwds = [str(u.get("password","")).lower(), str(u.get("password_alt","")).lower(), str(u.get("password_alt2","")).lower(), str(u.get("password_alt3","")).lower()]
                if pwd_lower not in stored_pwds and pwd_input not in [u.get("password",""), u.get("password_alt",""), u.get("password_alt2","")]:
                    # Also check exact case
                    if pwd_input != u.get("password","") and pwd_lower != str(u.get("password","")).lower():
                        continue
                # Check role
                if u["peranan"] in target_roles:
                    found = u
                    break
            
            if found:
                st.session_state.login = True
                st.session_state.user = found
                # v28: Clear cache lama supaya tak tunjuk BK001 bila login BK009
                st.session_state['current_ppd'] = found.get('ppd','-')
                st.session_state['current_sekolah'] = found.get('nama_pusat') or found.get('nama','-')
                st.session_state['current_no_pusat'] = found.get('no_pusat','-')
                st.success(f"✅ Selamat datang {found['nama']}!")
                st.balloons()
                st.rerun()
            else:
                st.error("❌ ID atau Kata Laluan salah. Sila semak peranan dan cuba lagi.")
                with st.expander("🔧 Fix: BA001 tak boleh login? Klik sini"):
                    st.write("App masih guna data lama KP001. Klik Reset untuk buat 504 ID baru (BA001, BK002...)")
                    if st.button("🔄 RESET DATABASE - Buat Semula 504 KP (BA001 etc)"):
                        if emergency_reset_users():
                            st.success("✅ Database direset! Sila Refresh dan cuba login BA001 / ba001 semula")
                            st.balloons()
                        else:
                            st.error("Gagal reset")
                # Debug login cleaned
        
        # PANDUAN DIBUANG IKUT REQUEST KASHAH
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Features bawah
        st.markdown('''
        <div style="margin-top:18px; display:grid; grid-template-columns:1fr 1fr; gap:10px;">
            <div class="feature-box">
                <div style="font-size:20px;">📝</div>
                <div style="font-size:11px; color:white; font-weight:600; margin-top:4px;">Laporan Pantas<br>KP Isi Terus</div>
            </div>
            <div class="feature-box">
                <div style="font-size:20px;">🖨️</div>
                <div style="font-size:11px; color:white; font-weight:600; margin-top:4px;">Cetak BLPP<br>Auto PDF</div>
            </div>
        </div>
        <div style="text-align:center; margin-top:20px; font-size:10px; color:rgba(255,255,255,0.6);">
            © 2026 Sektor Pentaksiran & Peperiksaan, JPN Selangor<br>Versi 4.1 FINAL • 504 Pusat • JPN Selangor • 2026
        </div>
        ''', unsafe_allow_html=True)
    
    st.stop()

# ========== DASHBOARD UTAMA ==========
user = st.session_state.user
peranan_user = user["peranan"]

st.markdown(f"""<div class="header-biru" style="text-align:center; padding:22px 10px;"><h1 style="font-size:22px; font-weight:800; margin:0; letter-spacing:0.5px;">📋 SISTEM LAPORAN PUSAT PEPERIKSAAN</h1><h2 style="font-size:14px; font-weight:500; margin:8px 0 0 0; color:#E3F2FD;">Sektor Pentaksiran dan Peperiksaan, Jabatan Pendidikan Selangor</h2><div style="font-size:12px; margin-top:10px; color:#FFD700; font-weight:700; letter-spacing:0.8px; text-shadow:0 1px 3px rgba(0,0,0,0.3);">Created by Akash | Peranan: {peranan_user}</div></div>""", unsafe_allow_html=True)

# ========== LAYOUT TANPA SIDEBAR - MENU KEKAL DI KIRI DALAM MAIN PAGE ==========
# Sembunyikan sidebar asal Streamlit sepenuhnya
st.markdown('''
<style>
section[data-testid="stSidebar"] {display: none !important;}
div[data-testid="stSidebarNav"] {display: none !important;}
</style>
''', unsafe_allow_html=True)

# Header kekal atas
# Buat 2 kolum: kiri untuk menu (kekalkan kedudukan macam sidebar), kanan untuk kandungan
col_menu, col_content = st.columns([1.1, 4], gap="medium")

with col_menu:
    # v28 FIX KASHAH: Untuk KP, guna user data terus, bukan session_state lama - elak BK001 tunjuk bila login BK009
    if peranan_user == "KP":
        current_ppd_display = user.get('ppd','-')
        current_sekolah_display = user.get('nama_pusat') or user.get('nama','-')
        current_no_pusat_display = user.get('no_pusat','-')
        # sync session untuk next render
        st.session_state['current_ppd'] = current_ppd_display
        st.session_state['current_sekolah'] = current_sekolah_display
        st.session_state['current_no_pusat'] = current_no_pusat_display
    else:
        current_ppd_display = st.session_state.get('current_ppd', user.get('ppd','-'))
        current_sekolah_display = st.session_state.get('current_sekolah', user.get('nama','-'))
        current_no_pusat_display = st.session_state.get('current_no_pusat', '-')
    
    st.markdown(f'''
    <div class="card" style="position:sticky; top:18px; padding:16px; border:2px solid #E2E8F0;">
        <div style="font-size:12px; font-weight:700; color:#0F172A; margin-bottom:8px; line-height:1.3;">{current_sekolah_display}</div>
        <div style="font-size:11px; color:#475569; margin:3px 0;"><b>No Pusat:</b> {current_no_pusat_display}</div>
        <div style="font-size:11px; background:#EEF2FF; color:#1E40AF; padding:6px 10px; border-radius:8px; display:block; margin-top:8px; font-weight:600;">PPD: {current_ppd_display}</div>
        <div style="font-size:10px; color:#94A3B8; margin-top:6px;">{peranan_user}</div>
        <hr style="margin:12px 0;">
    ''', unsafe_allow_html=True)
    
    # Menu ikut peranan - JPN ada Padam Laporan
    # FIX KASHAH: LP sama macam JPN cuma tiada Delete & Pengurusan Pengguna
    if peranan_user == "ADMIN_JPN":
        menu_items = ["📊 Dashboard", "📋 Senarai Laporan", "🖨️ Cetak Borang BLPP", "📦 Download Pukal ZIP", "🗑️ Padam Laporan", "👥 Pengurusan Pengguna (JPN)", "🚪 Log Keluar"]
    elif peranan_user == "PPD":
        menu_items = ["📊 Dashboard", "📋 Senarai Laporan", "🖨️ Cetak Borang BLPP", "📦 Download Pukal ZIP", "📈 Statistik PPD", "🚪 Log Keluar"]
    elif peranan_user == "LP":
        menu_items = ["📊 Dashboard", "📋 Senarai Laporan", "🖨️ Cetak Borang BLPP", "📦 Download Pukal ZIP", "🚪 Log Keluar"]
    else:
        menu_items = ["📝 Isi Borang BLPP (KP)", "🖨️ Cetak Semula Laporan Saya", "📋 Senarai Laporan", "📈 Statistik Saya", "🚪 Log Keluar"]
    menu = st.radio("Menu", menu_items, label_visibility="collapsed")
    
    st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
    if st.button("🚪 Log Keluar", use_container_width=True):
        st.session_state.login=False
        st.session_state.user=None
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# Semua kandungan selepas ini akan masuk dalam col_content
with col_content:


    reports = load_reports()

    # ========== MENU 1: ISI BORANG BLPP (KP) ==========
    if menu.startswith("📝") or (menu.startswith("📊") and peranan_user=="KP"):
        st.markdown('<div class="card"><h3>📝 Borang Laporan Pengendalian Peperiksaan - BLPP LP/Am 90</h3><p style="font-size:12px; color:#64748B;">Borang ini hendaklah digunakan semasa peperiksaan untuk melaporkan perkara luar dugaan (sakit tiba-tiba, kemalangan, ditahan polis dll). 504 Ketua Pengawas akan isi setiap hari selama 20 hari peperiksaan.</p></div>', unsafe_allow_html=True)
    
        if peranan_user != "KP":
            st.info("Hanya Ketua Pengawas (KP) boleh isi borang. PPD/JPN/LP boleh view laporan.")
        else:
            # v22 FIX KASHAH - KP tak perlu pilih pusat lagi, terus ikut data Excel Github
            # Load centers dari Excel real - untuk mapping yang betul
            try:
                centers_spm = load_spm_centers()
            except:
                centers_spm = []
            try:
                centers_amali = load_amali_centers()
            except:
                centers_amali = []
            centers_dict = {}
            for c in (centers_spm + centers_amali):
                centers_dict[c["No_Pusat"]] = c
            centers = list(centers_dict.values())
            exam_mode = "SPM"
            if not centers:
                try:
                    centers = get_all_pusat()
                except:
                    centers = []
            centers = ensure_centers(centers) if 'ensure_centers' in globals() else centers
            
            # v50 FIX: Paksa PPD ikut Kod prefix BC=Kuala Selangor, BJ=Sepang
            user_no_pusat = str(user.get("no_pusat","")).strip().upper()
            user_nama_pusat_excel = ""
            prefix_kod = user_no_pusat[:2].upper() if len(user_no_pusat)>=2 else ""
            forced_ppd = KOD_PPD_MAPPING.get(prefix_kod, user.get("ppd",""))
            user_ppd_excel = forced_ppd
            user_kod_excel = prefix_kod
            
            # Cari dalam Excel real
            found_in_excel = None
            for c in centers:
                if str(c.get("No_Pusat","")).strip().upper() == user_no_pusat:
                    found_in_excel = c
                    break
            if found_in_excel:
                user_nama_pusat_excel = found_in_excel.get("Nama_Pusat") or found_in_excel.get("Nama_Sekolah","")
                selected_sekolah = user_nama_pusat_excel
                auto_no_pusat = user_no_pusat
                auto_kod_ppd = prefix_kod
                selected_ppd = forced_ppd
                user_ppd_excel = forced_ppd
                user_kod_excel = prefix_kod
            else:
                # fallback kalau tak jumpa dalam Excel - guna data user
                selected_sekolah = user.get("nama_pusat", "")
                auto_no_pusat = user_no_pusat
                auto_kod_ppd = user_kod_excel
                selected_ppd = user.get("ppd","")
                st.caption(f"Pusat {user_no_pusat} - data login")
            
            # PAPARAN LOCKED - Tak ada dropdown lagi untuk KP
            st.markdown(f'''
            <div style="background:#f0fdf4; border:2px solid #22c55e; border-radius:12px; padding:16px; margin:12px 0;">
                <div style="font-weight:800; color:#15803d; font-size:14px; margin-bottom:8px;">✅ Pusat Anda (Locked - Ikut Data Excel)</div>
                <div style="display:grid; grid-template-columns:1fr 2fr 1fr; gap:12px; font-size:13px;">
                    <div><b>No. Pusat:</b><br><span style="font-size:16px; font-weight:800; color:#0f172a;">{auto_no_pusat}</span><br><span style="font-size:11px; color:#64748b;">Kod: {auto_kod_ppd} = {user_ppd_excel}</span></div>
                    <div><b>Nama Pusat:</b><br><span style="font-weight:700; color:#0f172a;">{selected_sekolah}</span></div>
                    <div><b>PPD:</b><br><span style="font-weight:600;">{user_ppd_excel}</span><br><span style="font-size:11px; color:#64748b;">ID: {user.get("id")}</span></div>
                </div>
                <div style="margin-top:8px; font-size:11px; color:#16a34a;">🔒 Data ini dari data_v1.1_3_1.xlsx - sama macam di Github. KP tidak boleh tukar pusat lain.</div>
            </div>
            ''', unsafe_allow_html=True)
            
            # Maklumat Pusat - RINGKAS & PROFESIONAL (header biru dibuang ikut request Kashah)
            # Tiada header besar, terus papar kotak edit sahaja
            
            # v22 - Final values - terus dari locked box, tak ada edit lagi
            final_no_pusat = auto_no_pusat
            final_nama_sekolah = selected_sekolah
            final_kod_ppd = auto_kod_ppd
            selected_ppd = user_ppd_excel
            
            # Hidden - untuk compatibility dengan kod bawah
            edit_no_pusat = final_no_pusat
            edit_nama_sekolah = final_nama_sekolah
            
            if not final_no_pusat or not final_nama_sekolah:
                st.info("👆 Sila pilih Daerah dan Nama Sekolah")
                st.stop()
            
            auto_no_pusat = final_no_pusat
            selected_sekolah = final_nama_sekolah
            auto_kod_ppd = final_kod_ppd
            
            # SYNC ke column kiri - simpan pilihan sebenar
            st.session_state['current_ppd'] = selected_ppd
            st.session_state['current_sekolah'] = selected_sekolah
            st.session_state['current_no_pusat'] = auto_no_pusat
            

        
            # Form BLPP sebenar - RINGKAS - FIX KASHAH: Peperiksaan & Sidang diluar form supaya dynamic
            st.markdown("---")
            st.markdown("#### 📝 Maklumat Peperiksaan (Pilih dulu sebelum isi kes)")
            
            colA, colB, colC = st.columns(3)
            with colA:
                try:
                    amali_idx = PEPERIKSAAN_LIST.index("Ujian Amali Sains SPM")
                except:
                    amali_idx = 2 if len(PEPERIKSAAN_LIST) > 2 else 0
                # Simpan peperiksaan dalam session_state supaya dynamic sidang berfungsi
                if 'selected_peperiksaan' not in st.session_state:
                    _base_list = get_active_peperiksaan_list()
                    st.session_state.selected_peperiksaan = _base_list[0 if exam_mode=="SPM" else (amali_idx if amali_idx < len(_base_list) else 0)]
                
                active_pep_list = get_active_peperiksaan_list()
                # Safety: kalau session ada STAM tapi dah freeze, reset ke SPM 2026
                if st.session_state.selected_peperiksaan not in active_pep_list:
                    st.session_state.selected_peperiksaan = active_pep_list[0]
                peperiksaan = st.selectbox("Peperiksaan", active_pep_list, index=active_pep_list.index(st.session_state.selected_peperiksaan) if st.session_state.selected_peperiksaan in active_pep_list else 0, key="peperiksaan_select")
                st.session_state.selected_peperiksaan = peperiksaan
                tahun = st.text_input("Tahun", value="2026")
                no_pusat = st.text_input("No. Pusat (Auto dari Data)", value=auto_no_pusat, disabled=False)
            with colB:
                nama_pusat = st.text_input("Nama Pusat Peperiksaan (Auto)", value=selected_sekolah)
                ppd = st.text_input("PPD / Daerah (Auto - Locked to PPD Anda)", value=user["ppd"], disabled=True, help="PPD auto ikut login anda, tidak boleh tukar untuk pastikan laporan masuk ke PPD anda")
                kod_ppd_display = st.text_input("Kod PPD", value=auto_kod_ppd, disabled=True)
            with colC:
                tarikh = st.date_input("Tarikh Laporan", value=date.today())
                # FIX KASHAH: Sidang ikut jenis peperiksaan - SEKARANG DILUAR FORM JADI DYNAMIC
                if "Amali" in peperiksaan:
                    # Amali Sains - 4 sidang, 8pg-1tgh sahaja, tulis sidang berapa sahaja
                    sidang_options = [
                        "Sidang 1",
                        "Sidang 2", 
                        "Sidang 3",
                        "Sidang 4"
                    ]
                    sidang_help = "Amali Sains: 4 sidang, semua dari 8 pagi hingga 1 tengah hari - tulis sidang berapa sahaja"
                else:
                    # SPM I - 3 sidang dengan masa betul
                    sidang_options = [
                        "Sidang 1 (8.00 - 10.00 pagi)",
                        "Sidang 2 (10.00 pagi - 1.00 tengah hari)",
                        "Sidang 3 (2.00 - 5.00 petang)"
                    ]
                    sidang_help = "SPM: Sidang 1=8-10pg, Sidang 2=10pg-1tgh, Sidang 3=2-5ptg"
                
                sidang = st.selectbox("Sidang", sidang_options, help=sidang_help)
                nama_kp = st.text_input("Nama KP", value="")
                st.text_input("Jenis Peperiksaan (System)", value=exam_mode, disabled=True, key="exam_mode_hidden")
        
            # ===== FIX CRITICAL KASHAH: CHECKBOX & BILANGAN KES DI LUAR FORM - ADA BUTANG ENTER SENDIRI =====
            st.markdown("---")
            st.markdown("#### 📋 Butiran Kes Luar Dugaan")
            
            # Init session
            if "checkbox_tiada_kes" not in st.session_state:
                st.session_state["checkbox_tiada_kes"] = False
            if "num_kes_input" not in st.session_state:
                st.session_state["num_kes_input"] = 1
            
            col_check, col_enter = st.columns([3,1])
            with col_check:
                tiada_kes = st.checkbox("☑️ TIADA KES LUAR DUGAAN PADA HARI INI (Tick jika tiada kes)", value=st.session_state["checkbox_tiada_kes"], key="checkbox_tiada_kes")
            with col_enter:
                if st.button("↩️ ENTER / Kemaskini", key="btn_enter_kes", help="Tekan ENTER untuk kemaskini paparan tanpa hantar laporan"):
                    st.rerun()
            
            st.session_state["current_tiada_kes"] = tiada_kes
            
            kes_data = []
            if not tiada_kes:
                col_num, col_enter2 = st.columns([3,1])
                with col_num:
                    num_kes = st.number_input("Bilangan Kes (1-20) - Jika 2 calon, pilih 2", min_value=1, max_value=20, value=st.session_state.get("num_kes_input", 1), key="num_kes_input")
                with col_enter2:
                    if st.button("↩️ ENTER", key="btn_enter_num", help="Tekan ENTER untuk tambah Kes 2 tanpa hantar"):
                        st.rerun()
                
                st.info("💡 Lepas tick/untick atau tukar bilangan kes, tekan butang ↩️ ENTER untuk kemaskini paparan. Bukan butang Hantar Laporan!")
                
                for i in range(num_kes):
                    with st.container(border=True):
                        st.markdown(f"### 📋 Kes {i+1} - Calon Ke-{i+1}")
                        if i==1:
                            st.warning(f"📌 Kes {i+1} - Ini kes kedua! Terus muncul lepas tekan ENTER!")
                        c1, c2 = st.columns(2)
                        with c1:
                            ag = st.text_input(f"Angka Giliran Calon {i+1}", key=f"ag_{i}", placeholder="BB057001")
                        with c2:
                            ic = st.text_input(f"No. IC Calon {i+1}", key=f"ic_{i}", placeholder="990101-10-1234")
                        c4, c5 = st.columns(2)
                        with c4:
                            kes = st.text_area(f"📝 Butiran Kes {i+1}", key=f"kes_{i}", height=100, placeholder="Sakit tiba-tiba / kemalangan...")
                        with c5:
                            tindakan = st.text_area(f"✅ Tindakan Kes {i+1}", key=f"tind_{i}", height=100, placeholder="Tindakan diambil...")
                        # Simpan untuk submit nanti
                        st.session_state[f"ag_{i}_val"] = ag
                        st.session_state[f"ic_{i}_val"] = ic
                        st.session_state[f"kes_{i}_val"] = kes
                        st.session_state[f"tind_{i}_val"] = tindakan
            else:
                st.success("✅ TIADA KES - Tick bermakna tiada kes. Tekan ENTER untuk sahkan, kemudian boleh terus Hantar Laporan.")
                num_kes = 0
            
            # Form hanya untuk tandatangan & hantar - TIDAK ADA checkbox dalam form
            with st.form("blpp_form", clear_on_submit=True):
            
                st.markdown("---")
                st.markdown("### ✍️ Tandatangan Digital & Pengesahan Pelapor")
                colT1, colT2 = st.columns([1,1])
                with colT1:
                    pelapor_nama = st.text_input("Nama Pelapor (Ketua Pengawas)", value=user.get("nama_kp","") or "", placeholder="Contoh: AKASHAH BIN ISMAIL")
                    telefon_kp = st.text_input("No. Telefon KP (untuk WhatsApp PPD)", placeholder="012-3456789", value=user.get("telefon",""))
                    st.markdown("**Muat Naik Tandatangan Digital**")
                    st.caption("KP boleh upload gambar tandatangan atau lukis tandatangan")
                    tandatangan_file = st.file_uploader("Upload Tandatangan (PNG/JPG)", type=["png","jpg","jpeg"], key="tt_file")
                    use_camera = st.checkbox("📷 Guna Kamera untuk Tangkap Tandatangan")
                    camera_tt = None
                    if use_camera:
                        camera_tt = st.camera_input("Tangkap Gambar Tandatangan", key="cam_tt")
            
                with colT2:
                    st.markdown("**Preview Tandatangan**")
                    def clear_bg_signature(pil_img):
                        # Buang background - jadikan transparent, hanya tinggal dakwat hitam
                        try:
                            # Convert ke RGBA
                            img_rgba = pil_img.convert("RGBA")
                            # Resize kecil untuk process laju
                            datas = img_rgba.getdata()
                            new_data = []
                            for item in datas:
                                r,g,b,a = item
                                # Jika pixel gelap (dakwat tandatangan < 90) - kekalkan hitam
                                # Jika cerah (background kertas, bayang hitam tepi) - jadikan transparent
                                # Threshold 100 untuk bezakan dakwat dengan background
                                if r < 100 and g < 100 and b < 100:
                                    # Dakwat - jadikan hitam pekat opaque
                                    new_data.append((0, 0, 0, 255))
                                else:
                                    # Background - transparent
                                    new_data.append((255, 255, 255, 0))
                            img_rgba.putdata(new_data)
                            return img_rgba
                        except Exception as e:
                            return pil_img.convert("RGBA")
                    
                    tt_b64 = ""
                    if tandatangan_file:
                        img = Image.open(tandatangan_file)
                        img.thumbnail((600, 400))
                        # Clear background terus
                        img_cleared = clear_bg_signature(img)
                        buf = BytesIO()
                        img_cleared.save(buf, format="PNG", optimize=True)
                        tt_b64 = base64.b64encode(buf.getvalue()).decode()
                        col_prev1, col_prev2 = st.columns(2)
                        with col_prev1:
                            st.image(img, caption="Asal", width=200)
                        with col_prev2:
                            st.image(img_cleared, caption="Background Cleared (Transparent)", width=200)
                        st.success(f"✅ Tandatangan cleared: {len(tt_b64)/1024:.1f}KB - Background transparent")
                    elif camera_tt:
                        img = Image.open(camera_tt)
                        img.thumbnail((600, 400))
                        img_cleared = clear_bg_signature(img)
                        buf = BytesIO()
                        img_cleared.save(buf, format="PNG", optimize=True)
                        tt_b64 = base64.b64encode(buf.getvalue()).decode()
                        col_prev1, col_prev2 = st.columns(2)
                        with col_prev1:
                            st.image(img, caption="Asal Kamera", width=200)
                        with col_prev2:
                            st.image(img_cleared, caption="Cleared - Boleh letak belakang text", width=200)
                        st.success(f"✅ Camera cleared: {len(tt_b64)/1024:.1f}KB")
                    else:
                        st.info("Belum ada tandatangan. Sila upload atau guna kamera.")
                        # Lukis manual guna canvas simple - guna text input sebagai alternatif
                        st.markdown("Atau taip nama sebagai tandatangan digital:")
                        tt_text = st.text_input("Tandatangan Text", value="", key="tt_text")
                        if tt_text:
                            # Generate signature image - HITAM BOLD BESAR, bukan biru kecil
                            from PIL import ImageDraw, ImageFont
                            # Canvas lebih besar untuk signature jelas
                            sig_img = Image.new('RGBA', (500, 120), color=(255,255,255,0))
                            d = ImageDraw.Draw(sig_img)
                            # Cuba font bold, jika tiada guna default tapi besar
                            try:
                                # Cuba guna font yang lebih tebal
                                font_large = ImageFont.truetype("DejaVuSans-Bold.ttf", 36)
                            except:
                                try:
                                    font_large = ImageFont.load_default()
                                except:
                                    font_large = None
                            # Tulis nama dengan hitam pekat, besar
                            if font_large:
                                d.text((10, 30), tt_text, fill=(0,0,0,255), font=font_large)
                            else:
                                d.text((10, 30), tt_text, fill=(0,0,0,255))
                            # Convert ke RGBA dengan background transparent
                            buf = BytesIO()
                            sig_img.save(buf, format="PNG", optimize=True)
                            tt_b64 = base64.b64encode(buf.getvalue()).decode()
                            st.image(sig_img, width=300, caption=f"Signature: {tt_text}")
                            # Juga set pelapor_nama preview
                            st.info(f"Nama Pelapor akan jadi: {tt_text.upper()}")

            
                submitted = st.form_submit_button("📤 Hantar Laporan BLPP", type="primary", use_container_width=True)
            
                if submitted:
                    if not no_pusat or not nama_pusat:
                        st.error("Sila isi No. Pusat dan Nama Pusat!")
                    else:
                        # FIX: Jika Ketua Pengawas isi Tandatangan Text, guna itu sebagai Nama Pelapor
                        # Jika tt_text ada dan bukan nama sekolah default, override pelapor_nama
                        try:
                            tt_text_val = st.session_state.get('tt_text','') or tt_text
                        except:
                            tt_text_val = tt_text if 'tt_text' in locals() else ''
                        
                        # Jika KP isi Tandatangan Text, guna itu sebagai nama pelapor
                        try:
                            tt_text_val = tt_text_val or tt_text or ""
                        except:
                            tt_text_val = ""
                        
                        if tt_text_val and tt_text_val.strip() != "":
                            # Jika ada tandatangan text, guna itu
                            if pelapor_nama.strip() == "" or "SEKOLAH" in pelapor_nama.upper():
                                pelapor_nama = tt_text_val.strip().upper()
                        
                        # BUANG NAMA SEKOLAH - jika pelapor_nama masih ada SEKOLAH, kosongkan
                        if "SEKOLAH" in pelapor_nama.upper():
                            if tt_text_val and tt_text_val.strip() != "":
                                pelapor_nama = tt_text_val.strip().upper()
                            else:
                                pelapor_nama = ""  # Kosongkan saja jika KP tak isi apa-apa
                        
                        new_report = {
                            "id": f"BLPP-{datetime.now().strftime('%Y%m%d%H%M%S')}-{no_pusat}",
                            "peperiksaan": peperiksaan.upper().strip(),  # FIX: HURUF BESAR dalam report
                            "tahun": tahun,
                            "no_pusat": no_pusat,
                            "nama_pusat": nama_pusat,
                            "nama_kp": nama_kp,
                            "ppd": KOD_PPD_MAPPING.get(str(auto_kod_ppd).strip().upper(), user.get("ppd","")) if 'auto_kod_ppd' in locals() and auto_kod_ppd else user["ppd"],
                            "kod_ppd": str(auto_kod_ppd).strip().upper() if 'auto_kod_ppd' in locals() and auto_kod_ppd else PPD_TO_KOD.get(user.get("ppd",""),""),
                            "tarikh": tarikh.strftime("%Y-%m-%d"),
                            "sidang": sidang,
                            "tarikh_submit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "kes_data": [] if st.session_state.get("checkbox_tiada_kes", False) else [ {"angka_giliran": st.session_state.get(f"ag_{i}", "").strip(), "no_ic": st.session_state.get(f"ic_{i}", "").strip(), "kes": st.session_state.get(f"kes_{i}", "").strip(), "tindakan": st.session_state.get(f"tind_{i}", "").strip()} for i in range(st.session_state.get("num_kes_input", 1)) if st.session_state.get(f"ag_{i}", "").strip() or st.session_state.get(f"ic_{i}", "").strip() or st.session_state.get(f"kes_{i}", "").strip() or st.session_state.get(f"tind_{i}", "").strip() ] or [{"angka_giliran": "", "no_ic": "", "kes": "", "tindakan": ""}],
                            "tiada_kes": st.session_state.get("checkbox_tiada_kes", False),
                            "pelapor": pelapor_nama,
                            "id_kp": user["id"],
                            "telefon_kp": telefon_kp,
                            "tandatangan_b64": tt_b64,
                            "tandatangan_text": pelapor_nama
                        }
                        reports.append(new_report)
                        save_reports(reports)
                        if GSHEETS_AVAILABLE:
                            try:
                                if save_to_gsheets(new_report):
                                    st.success("✅ Juga disimpan ke Google Sheets!")
                            except:
                                pass
                        # Simpan ke session_state untuk paparan di luar form
                        st.session_state['last_report'] = new_report
                        st.session_state['last_submitted'] = True
                        st.success(f"✅ Laporan BLPP berjaya dihantar! ID: {new_report['id']}")
                        st.balloons()
                        st.info("⬇️ Scroll ke bawah untuk lihat borang lengkap berisi data, download PDF & hantar WhatsApp")

            # ===== PAPARAN SELEPAS HANTAR - DI LUAR FORM (FIX ERROR download_button in form) =====
            if st.session_state.get('last_submitted') and st.session_state.get('last_report'):
                new_report = st.session_state['last_report']
                no_pusat = new_report.get('no_pusat','')
                nama_pusat = new_report.get('nama_pusat','')
                peperiksaan = new_report.get('peperiksaan','')
                tahun = new_report.get('tahun','')
                ppd = new_report.get('ppd','')
                nama_kp = new_report.get('nama_kp','')
                telefon_kp = new_report.get('telefon_kp','')
                pelapor_nama = new_report.get('pelapor','')
                kes_data = new_report.get('kes_data',[])
                tiada_kes = new_report.get('tiada_kes',False)
                try:
                    tarikh_obj = datetime.strptime(new_report.get('tarikh',''), "%Y-%m-%d")
                except:
                    tarikh_obj = datetime.now()
            
                st.markdown("---")
                st.markdown("### 📄 Borang BLPP Lengkap")
                st.success("✅ Borang di bawah telah siap diisi dengan data KP - ini yang akan dihantar ke PPD!")
            
                html_filled = jana_blpp_html(new_report)
                st.components.v1.html(html_filled, height=950, scrolling=True)
            
                # PDF - di luar form, jadi download_button dibenarkan
                pdf_bytes = b""
                try:
                    pdf_bytes = jana_blpp_pdf(new_report)
                except Exception as e:
                    st.warning(f"PDF error: {e}")
                    pdf_bytes = b""
            
                if pdf_bytes and len(pdf_bytes) > 500:
                    st.download_button(
                        "📥 Download Borang BLPP Lengkap (PDF Berisi Data) - Untuk Hantar WhatsApp",
                        data=pdf_bytes,
                        file_name="BLPP_%s_%s_%s.pdf" % (no_pusat, tarikh_obj.strftime('%Y%m%d'), new_report['id']),
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True,
                        key="dl_pdf_fixed"
                    )
                    st.caption("✅ PDF %.1fKB berjaya dijana dengan data lengkap" % (len(pdf_bytes)/1024))
                else:
                    st.warning("PDF gagal dijana, guna HTML download")
                    st.download_button(
                        "📥 Download Borang BLPP (HTML Lengkap)",
                        data=html_filled.encode('utf-8'),
                        file_name="BLPP_%s_%s.html" % (no_pusat, tarikh_obj.strftime('%Y%m%d')),
                        mime="text/html",
                        use_container_width=True,
                        key="dl_html_fixed"
                    )
            
                st.markdown("---")
                st.markdown("### 📱 Hantar WhatsApp ke PPD - DENGAN DATA LENGKAP")
                st.info("💡 WhatsApp akan hantar mesej lengkap berisi semua data BLPP. Untuk hantar PDF, download PDF di atas dahulu, kemudian attach dalam WhatsApp secara manual")
            
                users_list = load_users()
                ppd_user = next((u for u in users_list if u["peranan"]=="PPD" and u["ppd"]==ppd), None)
                ppd_tel = ppd_user.get("telefon","") if ppd_user else ""
            
                # Input telefon PPD - di luar form
                ppd_tel_input = st.text_input("No. Telefon PPD (untuk WhatsApp)", placeholder="60123456789", value=ppd_tel if ppd_tel else "60123456789", key="ppd_tel_input_fixed")
            
                if tiada_kes:
                    wa_text = "*BORANG LAPORAN PENGENDALIAN PEPERIKSAAN (BLPP)*\n"
                    wa_text += "*LP/Am 90 Pin. 1/2023*\n\n"
                    wa_text += "*PEPERIKSAAN:* %s\n" % peperiksaan
                    wa_text += "*TAHUN:* %s\n" % tahun
                    wa_text += "*NO. PUSAT:* %s\n" % no_pusat
                    wa_text += "*NAMA PUSAT:* %s\n" % nama_pusat
                    wa_text += "*PPD:* %s\n" % ppd
                    wa_text += "*TARIKH:* %s\n" % tarikh_obj.strftime('%d/%m/%Y')
                    wa_text += "*NAMA KP:* %s\n" % nama_kp
                    wa_text += "*TEL KP:* %s\n\n" % telefon_kp
                    wa_text += "━━━━━━━━━━━━━━━━━━━━\n✅ *TIADA KES LUAR DUGAAN*\n━━━━━━━━━━━━━━━━━━━━\n\n"
                    wa_text += "*Pelapor:* %s\n" % pelapor_nama
                    wa_text += "*ID:* %s\n" % new_report['id']
                    wa_text += "*Hantar:* %s\n\n" % new_report['tarikh_submit']
                    wa_text += "_Sistem BLPP JPN Selangor - 504 Pusat_"
                else:
                    wa_text = "🚨 *BORANG LAPORAN PENGENDALIAN PEPERIKSAAN (BLPP)*\n"
                    wa_text += "*LP/Am 90 Pin. 1/2023 - ADA KES*\n\n"
                    wa_text += "*PEPERIKSAAN:* %s\n" % peperiksaan
                    wa_text += "*TAHUN:* %s\n" % tahun
                    wa_text += "*NO. PUSAT:* %s\n" % no_pusat
                    wa_text += "*NAMA PUSAT:* %s\n" % nama_pusat
                    wa_text += "*PPD:* %s\n" % ppd
                    wa_text += "*TARIKH:* %s\n" % tarikh_obj.strftime('%d/%m/%Y')
                    wa_text += "*NAMA KP:* %s\n" % nama_kp
                    wa_text += "*TEL KP:* %s\n\n" % telefon_kp
                    wa_text += "━━━━━━━━━━━━━━━━━━━━\n🚨 *ADA %s KES:*\n━━━━━━━━━━━━━━━━━━━━\n" % len(kes_data)
                    for i, k in enumerate(kes_data, 1):
                        ag = k.get('angka_giliran','-') or '-'
                        ic = k.get('no_ic','-') or '-'
                        kes = k.get('kes','-') or '-'
                        tind = k.get('tindakan','-') or '-'
                        wa_text += "\n*%s. ANGKA GILIRAN:* %s\n" % (i, ag)
                        wa_text += "   *NO IC:* %s\n" % ic
                        wa_text += "   *KES:* %s\n" % kes
                        wa_text += "   *TINDAKAN:* %s\n" % tind
                    wa_text += "\n━━━━━━━━━━━━━━━━━━━━\n"
                    wa_text += "*Pelapor:* %s\n" % pelapor_nama
                    wa_text += "*ID:* %s\n" % new_report['id']
                    wa_text += "*Hantar:* %s\n\n" % new_report['tarikh_submit']
                    wa_text += "_Sila ambil tindakan. PDF lengkap boleh download dari sistem._\n"
                    wa_text += "_Sistem BLPP JPN Selangor_"
            
                import urllib.parse
                wa_encoded = urllib.parse.quote(wa_text)
                clean_tel = ppd_tel_input.replace("-","").replace(" ","").strip()
                if clean_tel.startswith("0"):
                    clean_tel = "6" + clean_tel
                wa_link = "https://wa.me/%s?text=%s" % (clean_tel, wa_encoded) if clean_tel else "https://wa.me/?text=%s" % wa_encoded
            
                st.markdown("**Preview Mesej WhatsApp Lengkap (Berisi Data):**")
                st.text_area("Mesej akan dihantar (sudah berisi data):", value=wa_text, height=350, key="wa_preview_fixed_final")
            
                col_wa1, col_wa2 = st.columns(2)
                with col_wa1:
                    st.link_button("📱 Hantar WhatsApp ke PPD %s (Berisi Data)" % ppd, wa_link, use_container_width=True, type="primary")
                with col_wa2:
                    wa_link_general = "https://wa.me/?text=%s" % wa_encoded
                    st.link_button("📤 Share WhatsApp (Contact Lain)", wa_link_general, use_container_width=True)
            
                st.warning("⚠️ *Cara Hantar PDF Berisi Data:* 1) Download PDF di atas 2) Buka WhatsApp PPD 3) Attach File PDF 4) Paste mesej di atas")
            
                if st.button("🔄 Isi Laporan Baru", key="reset_last"):
                    st.session_state['last_submitted'] = False
                    st.session_state['last_report'] = None
                    st.rerun()
                    

            # Papar laporan saya (KP)
            st.markdown("---")
            st.markdown("#### 📜 Laporan Saya (20 Hari Peperiksaan) - LOCKED")
            # v21 FINAL FIX KASHAH: KP 002 tak boleh tengok laporan pusat lain - STRICT
            user_pusat_filter = str(user.get("no_pusat","")).strip().upper()
            user_id_filter = str(user.get("id","")).strip().upper()
            
            # v35 FINAL: Bersih - tak tunjuk debug info untuk user
            # st.caption(f"Pusat: {user_pusat_filter}")
            
            if user_pusat_filter and user_pusat_filter != "-" and user_pusat_filter != "":
                # STRICT: id_kp mesti sama DAN no_pusat mesti sama dengan pusat KP - FINAL CLEAN
                my_reports = [r for r in reports if str(r.get("id_kp","")).strip().upper()==user_id_filter and str(r.get("no_pusat","")).strip().upper()==user_pusat_filter]
                # v35 FINAL: Tak tunjuk LOCKED message, bersih
                # st.caption(f"{len(my_reports)} laporan untuk {user_pusat_filter}")
            else:
                my_reports = [r for r in reports if str(r.get("id_kp","")).strip().upper()==user_id_filter]
                st.caption(f"Menunjukkan laporan untuk ID: {user_id_filter}")
            
            if my_reports:
                df_my = pd.DataFrame(my_reports)
                # tambah info jumlah
                st.success(f"✅ {len(my_reports)} laporan untuk pusat {my_reports[0].get('no_pusat')} - {my_reports[0].get('nama_pusat','')}")
                cols_show = [c for c in ["id","tarikh","peperiksaan","no_pusat","nama_pusat","sidang","tiada_kes"] if c in df_my.columns]
                st.dataframe(df_my[cols_show].sort_values("tarikh", ascending=False), use_container_width=True)
                # download my reports
                output = df_my.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Download Laporan Saya (CSV) - Hanya BK002", data=output, file_name=f"Laporan_{user_pusat_filter}_{user['id']}.csv", mime="text/csv")
            else:
                if user_pusat_filter and user_pusat_filter != "-":
                    st.info(f"Belum ada laporan untuk {user_pusat_filter}. Sila isi borang di atas. Laporan anda akan kekal untuk {user_pusat_filter} sahaja.")
                else:
                    st.info("Belum ada laporan. Sila isi borang di atas untuk hari ini.")

    # ========== MENU 1B: CETAK SEMULA LAPORAN SAYA (KP) - v40 KASHAH: KP lupa cetak semasa buat laporan, nak cetak semula ==========
    elif menu.startswith("🖨️") and peranan_user == "KP":
        # Check if this is Cetak Semula Laporan Saya (KP) vs Cetak Borang BLPP (PPD/JPN/LP)
        if "Semula" in menu or "Saya" in menu:
            st.markdown('<div class="card"><h3>🖨️ Cetak Semula Laporan Saya (KP)</h3><p style="font-size:12px; color:#475569;">KP terlupa cetak semasa hantar laporan? Pilih laporan lama anda di bawah untuk cetak semula PDF berlogo. Semua laporan anda ikut No Pusat anda sahaja.</p></div>', unsafe_allow_html=True)
            
            # Filter laporan KP ini sahaja
            my_reports = [r for r in reports if str(r.get("id_kp","")).strip().upper() == str(user.get("id","")).strip().upper() and str(r.get("no_pusat","")).strip().upper() == str(user.get("no_pusat","")).strip().upper()]
            
            if not my_reports:
                st.warning("📭 Belum ada laporan anda. Sila isi Borang BLPP dahulu di menu Isi Borang BLPP.")
            else:
                st.success(f"📋 Anda ada {len(my_reports)} laporan untuk pusat {user.get('no_pusat')} - {user.get('nama_pusat','')}")
                
                # Filter ikut tarikh & peperiksaan
                from collections import Counter
                col_tarikh_kp, col_pep_kp = st.columns(2)
                with col_tarikh_kp:
                    all_tarikh_kp = sorted(list(set([r.get('tarikh','') for r in my_reports if r.get('tarikh')])), reverse=True)
                    tarikh_opts_kp = ["SEMUA TARIKH"] + all_tarikh_kp
                    sel_tarikh_kp = st.selectbox("Pilih Tarikh", tarikh_opts_kp, key="kp_cetak_tarikh_v40")
                with col_pep_kp:
                    all_pep_kp = sorted(list(set([str(r.get('peperiksaan','')).strip().upper() for r in my_reports if r.get('peperiksaan')])))
                    pep_opts_kp = ["SEMUA PEPERIKSAAN"] + all_pep_kp
                    sel_pep_kp = st.selectbox("Pilih Peperiksaan", pep_opts_kp, key="kp_cetak_pep_v40")
                
                filtered_kp = my_reports
                if sel_tarikh_kp != "SEMUA TARIKH":
                    filtered_kp = [r for r in filtered_kp if r.get('tarikh') == sel_tarikh_kp]
                if sel_pep_kp != "SEMUA PEPERIKSAAN":
                    filtered_kp = [r for r in filtered_kp if str(r.get('peperiksaan','')).strip().upper() == str(sel_pep_kp).strip().upper()]
                
                # Sort latest first
                filtered_kp = sorted(filtered_kp, key=lambda x: x.get('tarikh_submit',''), reverse=True)
                
                if not filtered_kp:
                    st.warning(f"Tiada laporan untuk Tarikh={sel_tarikh_kp}, Peperiksaan={sel_pep_kp}")
                else:
                    st.info(f"✅ {len(filtered_kp)} laporan ditemui - pilih satu untuk cetak semula")
                    
                    # Buat list pilihan dengan label cantik
                    report_map_kp = {}
                    for r in filtered_kp:
                        kes_label = "TIADA KES" if r.get('tiada_kes') else f"{len(r.get('kes_data',[]))} KES"
                        sidang_label = r.get('sidang','') or '-'
                        submit_time = r.get('tarikh_submit','') or r.get('tarikh','')
                        try:
                            from datetime import datetime as dt_parse
                            if ' ' in submit_time:
                                dt_obj = dt_parse.strptime(submit_time, "%Y-%m-%d %H:%M:%S")
                                masa_cantik = dt_obj.strftime("%d/%m/%Y %H:%M:%S")
                            else:
                                masa_cantik = submit_time
                        except:
                            masa_cantik = submit_time
                        label = f"{r.get('tarikh')} | {r.get('peperiksaan','')} | {sidang_label} | {kes_label} | Hantar: {masa_cantik} | ID: {r.get('id','')[:20]}"
                        report_map_kp[label] = r
                    
                    sel_label_kp = st.selectbox(f"📄 Pilih Laporan untuk Cetak Semula ({len(report_map_kp)} laporan)", list(report_map_kp.keys()), key="kp_cetak_select_v40")
                    
                    if sel_label_kp:
                        selected_report_kp = report_map_kp[sel_label_kp]
                        
                        # Preview info
                        col_info1, col_info2 = st.columns(2)
                        with col_info1:
                            st.markdown(f"""
                            <div style="background:#F8FAFF; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
                                <b>No Pusat:</b> {selected_report_kp.get('no_pusat')}<br>
                                <b>Nama Pusat:</b> {selected_report_kp.get('nama_pusat')}<br>
                                <b>Peperiksaan:</b> {selected_report_kp.get('peperiksaan')}<br>
                                <b>Tarikh:</b> {selected_report_kp.get('tarikh')}<br>
                                <b>Sidang:</b> {selected_report_kp.get('sidang','-')}
                            </div>
                            """, unsafe_allow_html=True)
                        with col_info2:
                            kes_count = 0 if selected_report_kp.get('tiada_kes') else len(selected_report_kp.get('kes_data',[]))
                            st.markdown(f"""
                            <div style="background:#F0FDF4; border:1px solid #BBF7D0; border-radius:10px; padding:12px;">
                                <b>Status:</b> {"TIADA KES" if selected_report_kp.get('tiada_kes') else f"{kes_count} KES"}<br>
                                <b>Hantar Pada:</b> {selected_report_kp.get('tarikh_submit','')}<br>
                                <b>Pelapor:</b> {selected_report_kp.get('pelapor','') or selected_report_kp.get('nama_kp','')}<br>
                                <b>ID Laporan:</b> {selected_report_kp.get('id','')[:30]}
                            </div>
                            """, unsafe_allow_html=True)
                        
                        st.markdown("---")
                        # Preview HTML & PDF
                        try:
                            html_preview = jana_blpp_html(selected_report_kp)
                            st.markdown("#### 👁️ Preview Borang BLPP")
                            st.components.v1.html(html_preview, height=800, scrolling=True)
                        except Exception as e:
                            st.error(f"Gagal preview: {e}")
                        
                        st.markdown("---")
                        col_pdf, col_html = st.columns(2)
                        with col_pdf:
                            try:
                                pdf_bytes = jana_blpp_pdf(selected_report_kp)
                                if isinstance(pdf_bytes, BytesIO):
                                    pdf_bytes = pdf_bytes.getvalue()
                                st.download_button(
                                    label="📥 Cetak Semula PDF Berlogo",
                                    data=pdf_bytes,
                                    file_name=f"BLPP_CETAK_SEMULA_{selected_report_kp.get('no_pusat')}_{selected_report_kp.get('tarikh')}_{selected_report_kp.get('id','')[:8]}.pdf",
                                    mime="application/pdf",
                                    key="kp_download_pdf_v40",
                                    type="primary",
                                    use_container_width=True
                                )
                            except Exception as e:
                                st.error(f"Gagal jana PDF: {e}")
                        with col_html:
                            try:
                                html_bytes = jana_blpp_html(selected_report_kp).encode('utf-8')
                                st.download_button(
                                    label="📄 Download HTML",
                                    data=html_bytes,
                                    file_name=f"BLPP_CETAK_SEMULA_{selected_report_kp.get('no_pusat')}_{selected_report_kp.get('tarikh')}.html",
                                    mime="text/html",
                                    key="kp_download_html_v40",
                                    use_container_width=True
                                )
                            except:
                                pass
        else:
            # Ini untuk PPD/JPN/LP yang guna menu 🖨️ Cetak Borang BLPP - akan di-handle di bawah
            pass

    # ========== MENU 1B: CETAK SEMULA LAPORAN SAYA (KP) - KASHAH: KP lupa download, boleh cetak semula ==========
    elif menu.startswith("🖨️") and peranan_user == "KP":
        st.markdown('<div class="card"><h3>🖨️ Cetak Semula Laporan Saya (KP)</h3><p style="font-size:12px; color:#475569;">Fungsi untuk KP yang terlupa download PDF masa hantar laporan. Pilih laporan lama anda di bawah untuk cetak semula. Hanya laporan ikut No Pusat anda sahaja.</p></div>', unsafe_allow_html=True)
        
        my_reports = [r for r in reports if str(r.get("id_kp","")).strip().upper() == str(user.get("id","")).strip().upper()]
        
        if not my_reports:
            st.warning("📭 Belum ada laporan anda. Sila isi Borang BLPP dahulu di menu Isi Borang.")
        else:
            st.success(f"📋 Jumpa {len(my_reports)} laporan anda - {user.get('no_pusat')} - {user.get('nama_pusat','')}")
            
            col1, col2 = st.columns(2)
            with col1:
                all_tarikh = sorted(list(set([r.get('tarikh','') for r in my_reports if r.get('tarikh')])), reverse=True)
                sel_tarikh = st.selectbox("Filter Tarikh", ["SEMUA TARIKH"] + all_tarikh, key="kp_cetak_tarikh")
            with col2:
                all_pep = sorted(list(set([str(r.get('peperiksaan','')).strip().upper() for r in my_reports if r.get('peperiksaan')])))
                sel_pep = st.selectbox("Filter Peperiksaan", ["SEMUA PEPERIKSAAN"] + all_pep, key="kp_cetak_pep")
            
            filtered = my_reports
            if sel_tarikh != "SEMUA TARIKH":
                filtered = [r for r in filtered if r.get('tarikh') == sel_tarikh]
            if sel_pep != "SEMUA PEPERIKSAAN":
                filtered = [r for r in filtered if str(r.get('peperiksaan','')).strip().upper() == sel_pep]
            
            filtered = sorted(filtered, key=lambda x: x.get('tarikh_submit',''), reverse=True)
            
            if not filtered:
                st.warning("Tiada laporan untuk filter tersebut.")
            else:
                report_map = {}
                for r in filtered:
                    kes_lbl = "TIADA KES" if r.get('tiada_kes') else f"{len(r.get('kes_data',[]))} KES"
                    label = f"{r.get('tarikh')} | {r.get('peperiksaan','')} | {kes_lbl} | Hantar: {str(r.get('tarikh_submit',''))[:19]}"
                    report_map[label] = r
                
                sel_label = st.selectbox(f"Pilih Laporan Untuk Cetak Semula ({len(report_map)} laporan)", list(report_map.keys()), key="kp_cetak_select")
                
                if sel_label:
                    sel_r = report_map[sel_label]
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.info(f"""**No Pusat:** {sel_r.get('no_pusat')}

**Nama Pusat:** {sel_r.get('nama_pusat')}

**Peperiksaan:** {sel_r.get('peperiksaan')}

**Tarikh:** {sel_r.get('tarikh')}""")
                    with c2:
                        kes_n = 0 if sel_r.get('tiada_kes') else len(sel_r.get('kes_data',[]))
                        st.success(f"""**Status:** {'TIADA KES' if sel_r.get('tiada_kes') else f'{kes_n} KES'}

**Hantar:** {sel_r.get('tarikh_submit','')}

**Pelapor:** {sel_r.get('pelapor','')}

**ID:** {sel_r.get('id','')[:25]}""")
                    
                    try:
                        html_preview = jana_blpp_html(sel_r)
                        st.markdown("#### 👁️ Preview Borang (Untuk Cetak Semula)")
                        st.components.v1.html(html_preview, height=800, scrolling=True)
                    except Exception as e:
                        st.error(f"Preview gagal: {e}")
                    
                    col_pdf, col_html = st.columns(2)
                    with col_pdf:
                        try:
                            pdf_bytes = jana_blpp_pdf(sel_r)
                            if isinstance(pdf_bytes, BytesIO):
                                pdf_bytes = pdf_bytes.getvalue()
                            st.download_button("📥 Download Semula PDF Berlogo", data=pdf_bytes, file_name=f"BLPP_{sel_r.get('no_pusat')}_{sel_r.get('tarikh')}_{sel_r.get('id','')[:8]}.pdf", mime="application/pdf", type="primary", use_container_width=True, key="kp_dl_pdf")
                        except Exception as e:
                            st.error(f"PDF gagal: {e}")
                    with col_html:
                        try:
                            html_b = jana_blpp_html(sel_r).encode('utf-8')
                            st.download_button("📄 Download HTML", data=html_b, file_name=f"BLPP_{sel_r.get('no_pusat')}_{sel_r.get('tarikh')}.html", mime="text/html", use_container_width=True, key="kp_dl_html")
                        except:
                            pass

    # ========== MENU 2: SENARAI LAPORAN ==========
    elif menu.startswith("📋"):
        st.markdown('<div class="card"><h3>📋 Senarai Laporan BLPP</h3></div>', unsafe_allow_html=True)
    
        # Filter
        colF1, colF2, colF3, colF4 = st.columns(4)
        with colF1:
            filter_ppd = st.selectbox("Pilih PPD", ["Semua"] + PPD_LIST, index=0 if peranan_user in ["ADMIN_JPN","LP"] else PPD_LIST.index(user["ppd"])+1 if user["ppd"] in PPD_LIST else 0)
        with colF2:
            filter_tarikh = st.date_input("Tarikh (optional)", value=None)
        with colF3:
            filter_pep = st.selectbox("Peperiksaan", ["Semua"] + PEPERIKSAAN_LIST)
        with colF4:
            filter_kes = st.selectbox("Jenis", ["Semua", "Ada Kes", "Tiada Kes"])

        for r in reports:
            kod_f = str(r.get("kod_ppd","")).upper().strip()
            kod_p = str(r.get("no_pusat",""))[:2].upper()
            kod = kod_f if kod_f in KOD_PPD_MAPPING else kod_p
            if kod in KOD_PPD_MAPPING:
                r["ppd"] = KOD_PPD_MAPPING[kod]
                r["kod_ppd"] = kod
        filtered = reports

        # v20 FIX KASHAH: KP lihat laporan sendiri sahaja - ikut No Pusat sendiri
        if peranan_user == "KP":
            kp_pusat = str(user.get("no_pusat","")).strip().upper()
            if kp_pusat and kp_pusat != "-":
                filtered = [r for r in filtered if str(r.get("no_pusat","")).strip().upper() == kp_pusat and r.get("id_kp") == user["id"]]
            else:
                filtered = [r for r in filtered if r.get("id_kp") == user["id"]]
        if peranan_user == "PPD":
            user_ppd = user["ppd"]
            # v29 STRICT: PPD Klang hanya nampak BA, tak campur BB/BC etc - ikut No Pusat prefix sahaja
            filtered = [r for r in filtered if is_report_for_ppd_strict(r, user_ppd)]
        elif peranan_user == "KP":
            filtered = [r for r in filtered if r.get("id_kp")==user["id"]]
    
        if filter_ppd != "Semua":
            # v31 STRICT: Filter Senarai Laporan JPN ikut No Pusat prefix - tak campur
            filtered = [r for r in filtered if is_report_for_ppd_strict(r, filter_ppd)]
        if filter_tarikh:
            filtered = [r for r in filtered if r.get("tarikh")==filter_tarikh.strftime("%Y-%m-%d")]
        if filter_pep != "Semua":
            filtered = [r for r in filtered if r.get("peperiksaan")==filter_pep]
        if filter_kes == "Ada Kes":
            filtered = [r for r in filtered if not r.get("tiada_kes")]
        if filter_kes == "Tiada Kes":
            filtered = [r for r in filtered if r.get("tiada_kes")]

        st.markdown(f"**Jumlah Laporan: {len(filtered)}** (dari {len(reports)} keseluruhan) | 504 Pusat | 20 Hari Peperiksaan")
    
        if filtered:
            df = pd.DataFrame(filtered)
            # Show table
            display_cols = ["id","tarikh","peperiksaan","no_pusat","nama_pusat","ppd","nama_kp","tiada_kes"]
            existing_cols = [c for c in display_cols if c in df.columns]
            st.dataframe(df[existing_cols].sort_values("tarikh", ascending=False), use_container_width=True, height=400)
        
            # Download Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Muat Turun Excel", data=output.getvalue(), file_name=f"BLPP_Laporan_{filter_ppd}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Tiada laporan ditemui dengan filter tersebut.")

    # ========== MENU 3: CETAK BORANG BLPP - FIX UNTUK PPD (KASHAH) ==========
    elif menu.startswith("🖨️") and peranan_user != "KP":
        st.markdown('<div class="card"><h3>🖨️ Cetak Borang BLPP LP/Am 90 Pin. 1/2023 (Official) - PPD Enhanced</h3><p style="font-size:12px;">Proses: Login PPD → Pilih Daerah (PPD) → Pilih Tarikh → Pilih Sekolah/Pusat → Cetak. PPD Klang akan nampak semua laporan PPD Klang sahaja.</p></div>', unsafe_allow_html=True)
    
        if not reports:
            st.warning("Belum ada laporan untuk dicetak.")
        else:
            # v35 FINAL CLEAN - tak tunjuk debug total
            available_reports = reports
            
            if peranan_user == "PPD":
                st.info(f"📍 PPD {user['ppd']} - {len([r for r in reports if is_report_for_ppd_strict(r, user['ppd'])])} laporan daerah {user['ppd']}")
            
            elif peranan_user == "KP":
                available_reports = [r for r in reports if r.get("id_kp")==user["id"]]
        
            # LANGKAH 1: PILIH DAERAH/PPD - BUANG UNTUK PPD (KASHAH REQUEST)
            if peranan_user == "PPD":
                sel_ppd = user["ppd"]
                kod_ppd_user = user.get("kod_ppd","")
                allowed = get_kods_for_ppd(sel_ppd)
                st.markdown(f"<div style='background:#EEF2FF; padding:10px; border-radius:8px; border-left:4px solid #3B82F6;'><b>📍 PPD {sel_ppd} - {len([r for r in reports if is_report_for_ppd_strict(r, sel_ppd)])} laporan</b></div>", unsafe_allow_html=True)
                # v35 FINAL: STRICT by No Pusat prefix
                available_reports = [r for r in reports if is_report_for_ppd_strict(r, sel_ppd)]
                st.caption(f"✅ {len(available_reports)} laporan untuk PPD {sel_ppd} ditemui")
                if len(available_reports)==0 and len(reports)>0:
                    st.info(f"Belum ada laporan untuk {sel_ppd}. KP daerah {sel_ppd} belum hantar lagi.")
            else:
                # JPN/LP/Admin masih perlu pilih PPD
                sel_ppd = st.selectbox("1️⃣ Pilih Daerah (PPD)", ["Semua"] + PPD_LIST, help="JPN/LP boleh pilih mana-mana PPD")
                if sel_ppd != "Semua":
                    # FIX ROBUST - Match Petaling Utama juga via kod BK & no_pusat
                    def match_jpn_ppd(r):
                        r_ppd = str(r.get("ppd","")).strip()
                        r_kod = str(r.get("kod_ppd","") or r.get("no_pusat","")[:2]).strip().upper()
                        sel = sel_ppd.strip()
                        if r_ppd == sel:
                            return True
                        # Map kod to PPD
                        kod_sel = PPD_TO_KOD.get(sel, "")
                        if kod_sel and r_kod == kod_sel:
                            return True
                        if sel == "Petaling Utama" and (r_ppd in ["BK", "Petaling Utama"] or str(r.get("no_pusat","")).upper().startswith("BK")):
                            return True
                        if sel == "Petaling Perdana" and (r_ppd in ["BH", "Petaling Perdana"] or str(r.get("no_pusat","")).upper().startswith("BH")):
                            return True
                        # Generic: check if kod matches selected PPD
                        if PPD_TO_KOD.get(sel) == r_kod:
                            return True
                        return False
                    available_reports = [r for r in available_reports if match_jpn_ppd(r)]
        
            if not available_reports:
                st.warning(f"Tiada laporan untuk PPD {sel_ppd} - KP belum hantar laporan lagi")
                selected_report = None
            else:
                # LANGKAH 2: PILIH TARIKH - hanya tarikh yang ada laporan untuk PPD tersebut
                tarikh_list = sorted(list(set([r.get('tarikh','') for r in available_reports if r.get('tarikh')])), reverse=True)
                if not tarikh_list:
                    st.info("Tiada tarikh laporan ditemui")
                    selected_report = None
                else:
                    # Format tarikh cantik
                    tarikh_options = ["Semua Tarikh"] + tarikh_list
                    sel_tarikh = st.selectbox(f"2️⃣ Pilih Tarikh Laporan ({len(tarikh_list)} tarikh ada laporan)", tarikh_options, help="Hanya tarikh yang KP telah hantar laporan untuk PPD ini")
                    
                    filtered_by_tarikh = available_reports
                    if sel_tarikh != "Semua Tarikh":
                        filtered_by_tarikh = [r for r in available_reports if r.get('tarikh')==sel_tarikh]
                    
                    # LANGKAH 3: PILIH SEKOLAH/PUSAT - hanya pusat yang telah hantar laporan
                    pusat_map = {}
                    for r in filtered_by_tarikh:
                        key = f"{r.get('no_pusat','')} - {r.get('nama_pusat','')}"
                        # Simpan laporan terkini untuk setiap pusat (jika ada banyak tarikh sama, ambil latest)
                        if key not in pusat_map:
                            pusat_map[key] = r
                        else:
                            # Jika ada duplicate, ambil yang latest submit
                            if r.get('tarikh_submit','') > pusat_map[key].get('tarikh_submit',''):
                                pusat_map[key] = r
                    
                    pusat_list = sorted(list(pusat_map.keys()))
                    
                    if not pusat_list:
                        st.info("Tiada pusat untuk tarikh tersebut")
                        selected_report = None
                    else:
                        sel_pusat_key = st.selectbox(f"3️⃣ Pilih Pusat/Sekolah ({len(pusat_list)} pusat telah hantar laporan)", ["Semua Pusat"] + pusat_list, help="Hanya pusat yang KP telah hantar laporan untuk PPD dan tarikh ini")
                        
                        final_filtered = filtered_by_tarikh
                        if sel_pusat_key != "Semua Pusat":
                            # Filter ikut pusat
                            no_pusat_selected = sel_pusat_key.split(" - ")[0]
                            final_filtered = [r for r in filtered_by_tarikh if r.get('no_pusat')==no_pusat_selected]
                        
                        # LANGKAH 3B: PILIH SIDANG - FIX KASHAH 3 SIDANG
                        sidang_list = sorted(list(set([r.get('sidang','') for r in final_filtered if r.get('sidang')])))
                        if sidang_list:
                            sel_sidang = st.selectbox(f"3️⃣b Pilih Sidang ({len(sidang_list)} sidang ada laporan)", ["Semua Sidang"] + sidang_list, help="Pilih sidang - KP isi sidang berapa")
                            if sel_sidang != "Semua Sidang":
                                final_filtered = [r for r in final_filtered if r.get('sidang')==sel_sidang]
                        else:
                            # Fallback jika laporan lama tak ada sidang
                            st.caption("ℹ️ Laporan lama tiada info sidang - semua sidang dipaparkan")
                        
                        # Sort ikut masa hantar - latest dulu untuk PPD mudah pilih
                        final_filtered = sorted(final_filtered, key=lambda x: x.get('tarikh_submit',''), reverse=True)
                        
                        # Jika masih banyak laporan untuk pusat sama (tarikh berbeza), pilih latest atau bagi pilih
                        if len(final_filtered) == 0:
                            selected_report = None
                        elif len(final_filtered) == 1:
                            selected_report = final_filtered[0]
                            st.success(f"✅ 1 laporan ditemui: {selected_report.get('no_pusat')} - {selected_report.get('nama_pusat')} - {selected_report.get('tarikh')}")
                        else:
                            # Banyak laporan untuk pusat yang sama (tarikh berbeza atau hantar banyak kali)
                            st.info(f"📋 {len(final_filtered)} laporan ditemui untuk filter ini - pilih satu untuk cetak")
                            report_options = {}
                            for r in final_filtered:
                                kes_label = "TIADA KES" if r.get('tiada_kes') else f"{len(r.get('kes_data',[]))} KES"
                                sidang_label = r.get('sidang','') or 'Tiada Sidang'
                                # FIX: Tunjuk masa KP isi/hantar untuk PPD boleh pilih ikut masa
                                submit_time = r.get('tarikh_submit','') or r.get('tarikh','')
                                # Format masa cantik
                                try:
                                    from datetime import datetime as dt_parse
                                    if ' ' in submit_time:
                                        dt_obj = dt_parse.strptime(submit_time, "%Y-%m-%d %H:%M:%S")
                                        masa_cantik = dt_obj.strftime("%d/%m/%Y %H:%M:%S")
                                        masa_short = dt_obj.strftime("%H:%M:%S")
                                    else:
                                        masa_cantik = submit_time
                                        masa_short = ""
                                except:
                                    masa_cantik = submit_time
                                    masa_short = ""
                                
                                label = f"{r['tarikh']} | {sidang_label} | {masa_short} | {r['no_pusat']} | {r['nama_pusat']} | {kes_label} | Hantar: {masa_cantik}"
                                report_options[label] = r
                            selected_label = st.selectbox("4️⃣ Pilih Laporan Spesifik untuk Cetak", list(report_options.keys()))
                            selected_report = report_options[selected_label]
            
            if selected_report:
                st.markdown("---")
                st.markdown("#### 📄 Preview Borang Rasmi (100% Sama Seperti Word LP)")
                # Jana HTML rasmi
                html = jana_blpp_html(selected_report)
                st.components.v1.html(html, height=950, scrolling=True)
            
                # Jana PDF rasmi
                pdf_bytes = b""
                try:
                    pdf_bytes = jana_blpp_pdf(selected_report)
                except Exception as e:
                    st.warning(f"PDF error: {e}")
                    pdf_bytes = b""
            
                st.markdown("---")
                st.markdown("### 📥 Download Borang (PDF Rasmi - 100% Sama Seperti Word)")
            
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    if pdf_bytes and len(pdf_bytes) > 500:
                        st.download_button(
                            "📥 Download PDF Rasmi (Boleh Print)",
                            data=pdf_bytes,
                            file_name=f"BLPP_{selected_report.get('no_pusat','')}_{selected_report.get('tarikh','')}_{selected_report.get('id','')}.pdf",
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True,
                            key=f"dl_pdf_cetak_{selected_report.get('id','')}"
                        )
                        st.caption(f"✅ PDF {len(pdf_bytes)/1024:.1f}KB - Logo rasmi, layout 100% sama Word")
                    else:
                        st.error("PDF gagal dijana, sila guna HTML")
            
                with col_dl2:
                    st.download_button(
                        "📄 Download HTML Rasmi (Boleh Print dari Browser)",
                        data=html.encode('utf-8'),
                        file_name=f"BLPP_{selected_report.get('no_pusat','')}_{selected_report.get('tarikh','')}.html",
                        mime="text/html",
                        use_container_width=True,
                        key=f"dl_html_cetak_{selected_report.get('id','')}"
                    )
                    st.caption("Buka file HTML, tekan Ctrl+P untuk print")
            
                st.info("💡 **Cara Print:** 1) Download PDF di atas 2) Buka PDF 3) Tekan Ctrl+P / File > Print 4) Pilih printer. PDF ini 100% sama macam borang Word rasmi LP!")

                # WhatsApp untuk cetak menu juga
                st.markdown("---")
                st.markdown("### 📱 Hantar WhatsApp (Dari Menu Cetak)")
                tiada_kes = selected_report.get('tiada_kes', False)
                kes_data = selected_report.get('kes_data', [])
                no_pusat = selected_report.get('no_pusat','')
                nama_pusat = selected_report.get('nama_pusat','')
                peperiksaan = selected_report.get('peperiksaan','')
                tahun = selected_report.get('tahun','')
                ppd = selected_report.get('ppd','')
                nama_kp = selected_report.get('nama_kp','')
                telefon_kp = selected_report.get('telefon_kp','')
                pelapor_nama = selected_report.get('pelapor','')
                tarikh_str = selected_report.get('tarikh','')
            
                try:
                    from datetime import datetime as dt
                    tarikh_obj = dt.strptime(tarikh_str, "%Y-%m-%d")
                    tarikh_formatted = tarikh_obj.strftime('%d/%m/%Y')
                except:
                    tarikh_formatted = tarikh_str
                    tarikh_obj = dt.now() if 'dt' in locals() else None
            
                if tiada_kes:
                    wa_text_cetak = "*BORANG LAPORAN PENGENDALIAN PEPERIKSAAN (BLPP)*\n*LP/Am 90 Pin. 1/2023*\n\n"
                    wa_text_cetak += "*PEPERIKSAAN:* %s\n" % peperiksaan
                    wa_text_cetak += "*NO. PUSAT:* %s\n" % no_pusat
                    wa_text_cetak += "*NAMA PUSAT:* %s\n" % nama_pusat
                    wa_text_cetak += "*PPD:* %s\n" % ppd
                    wa_text_cetak += "*TARIKH:* %s\n" % tarikh_formatted
                    wa_text_cetak += "*NAMA KP:* %s\n\n" % nama_kp
                    wa_text_cetak += "✅ *TIADA KES*\n"
                    wa_text_cetak += "*ID:* %s" % selected_report.get('id','')
                else:
                    wa_text_cetak = "🚨 *BLPP - ADA %s KES*\n" % len(kes_data)
                    wa_text_cetak += "*%s - %s*\n" % (no_pusat, nama_pusat)
                    wa_text_cetak += "*PPD:* %s\n" % ppd
                    for i, k in enumerate(kes_data[:3], 1):
                        wa_text_cetak += "%s. %s - %s\n" % (i, k.get('angka_giliran',''), k.get('kes','')[:60])
                    wa_text_cetak += "*ID:* %s" % selected_report.get('id','')
            
                import urllib.parse
                wa_encoded_cetak = urllib.parse.quote(wa_text_cetak)
                st.text_area("Mesej WhatsApp:", value=wa_text_cetak, height=150, key=f"wa_cetak_{selected_report.get('id','')}")
                wa_link_cetak = f"https://wa.me/?text={wa_encoded_cetak}"
                st.link_button("📱 Hantar WhatsApp", wa_link_cetak, use_container_width=True, key=f"wa_btn_{selected_report.get('id','')}")


    # ========== MENU BARU: DOWNLOAD PUKAL ZIP - FIX KASHAH: JPN PILIH DAERAH + TARIKH, PPD HANYA DAERAH DIA ==========
    elif menu.startswith("📦"):
        st.markdown('<div class="card"><h3>📦 Download Pukal ZIP - PDF Berlogo</h3><p style="font-size:12px;">JPN/PPD boleh download semua laporan ikut tarikh dalam 1 ZIP file. Semua PDF ada logo LP, tulisan sampai sempadan, garis lurus, tiada jarak, 1 muka TIADA KES.</p></div>', unsafe_allow_html=True)
        
        # v29 STRICT PPD - PPD Klang jangan ada data campur PPD lain - FILTER STRICT BY NO PUSAT PREFIX
        if peranan_user == "PPD":
            user_ppd = user.get('ppd','')
            allowed_kods = get_kods_for_ppd(user_ppd)
            reports_for_download = [r for r in reports if is_report_for_ppd_strict(r, user_ppd)]
            st.success(f"📍 PPD {user_ppd} - {len(reports_for_download)} laporan")
        else:
            reports_for_download = reports
            # Breakdown untuk JPN
            if reports_for_download:
                from collections import Counter
                all_ppd_counter = Counter([KOD_PPD_MAPPING.get(str(r.get('kod_ppd','') or r.get('no_pusat','')[:2]).upper(), r.get('ppd','') or 'TIADA PPD') for r in reports_for_download])
                all_ppd_detail = ", ".join([f"{ppd}={count}" for ppd, count in all_ppd_counter.items()])
                st.info(f"👑 JPN/LP - semua laporan {len(reports_for_download)}: {all_ppd_detail}")
        
        if not reports_for_download:
            if peranan_user == "PPD":
                st.warning(f"Belum ada laporan untuk daerah {user.get('ppd')}. KP daerah {user.get('ppd')} belum hantar.")
                # Debug cleaned for production
            else:
                st.warning("Belum ada laporan dalam sistem.")
        else:
            st.success(f"📊 Total laporan: {len(reports_for_download)} {'(Daerah ' + user.get('ppd') + ' sahaja)' if peranan_user=='PPD' else '(Semua Daerah)'}")
            from collections import Counter
            tarikh_counts = Counter([r.get('tarikh','') for r in reports_for_download if r.get('tarikh')])
            
            # JPN: Pilihan Daerah + Tarikh + Peperiksaan (v37 - KASHAH: tak campur peperiksaan)
            if peranan_user == "ADMIN_JPN" or peranan_user == "LP":
                st.markdown("#### 🎯 Pilih Daerah, Tarikh & Peperiksaan (JPN/LP - Semua Daerah)")
                col_ppd, col_tarikh, col_pep = st.columns(3)
                with col_ppd:
                    # FIX KASHAH: Dropdown JPN mesti tunjuk SEMUA DAERAH, bukan hanya Klang sahaja - ambil dari PPD_LIST penuh
                    def get_ppd_betul_dl(r):
                        kod = str(r.get('kod_ppd','') or r.get('no_pusat','')[:2]).upper()
                        return KOD_PPD_MAPPING.get(kod, r.get('ppd',''))
                    all_ppds_with_data = sorted(list(set([get_ppd_betul_dl(r) for r in reports_for_download if get_ppd_betul_dl(r)])))
                    ppd_options = ["SEMUA DAERAH"] + PPD_LIST
                    selected_ppd_filter = st.selectbox("Pilih Daerah / PPD", ppd_options, key="jpn_ppd_filter_v2")
                    from collections import Counter
                    ppd_count_map = Counter([get_ppd_betul_dl(r) for r in reports_for_download])
                    if selected_ppd_filter != "SEMUA DAERAH":
                        count_sel = ppd_count_map.get(selected_ppd_filter, 0)
                        if count_sel == 0:
                            st.warning(f"⚠️ Belum ada laporan di daerah {selected_ppd_filter}")
                with col_tarikh:
                    all_tarikhs = sorted(list(set([r.get('tarikh','') for r in reports_for_download if r.get('tarikh')])), reverse=True)
                    tarikh_options = ["SEMUA TARIKH"] + all_tarikhs
                    selected_tarikh_filter = st.selectbox("Pilih Tarikh", tarikh_options, key="jpn_tarikh_filter_v2")
                with col_pep:
                    # v37: Tambah pilihan peperiksaan - supaya tak campur SPM dengan Amali, Bertutur etc
                    all_peps = sorted(list(set([str(r.get('peperiksaan','')).strip().upper() for r in reports_for_download if r.get('peperiksaan')])))
                    # Guna PEPERIKSAAN_LIST asal + yang ada dalam data
                    pep_options = ["SEMUA PEPERIKSAAN"] + (PEPERIKSAAN_LIST if 'PEPERIKSAAN_LIST' in globals() else all_peps) 
                    # Pastikan unique dan uppercase
                    seen = set()
                    pep_options_unique = []
                    for p in pep_options:
                        pu = str(p).strip().upper()
                        if pu not in seen:
                            seen.add(pu)
                            pep_options_unique.append(p)
                    # Jika data ada peperiksaan yang tak ada dalam list, tambah
                    for ap in all_peps:
                        if ap not in seen:
                            pep_options_unique.append(ap)
                    selected_pep_filter = st.selectbox("Pilih Peperiksaan", pep_options_unique, key="jpn_pep_filter_v37")
                
                filtered_jpn = reports_for_download
                if selected_ppd_filter != "SEMUA DAERAH":
                    # v31 STRICT JPN Download
                    filtered_jpn = [r for r in filtered_jpn if is_report_for_ppd_strict(r, selected_ppd_filter)]
                if selected_tarikh_filter != "SEMUA TARIKH":
                    filtered_jpn = [r for r in filtered_jpn if r.get('tarikh') == selected_tarikh_filter]
                if selected_pep_filter != "SEMUA PEPERIKSAAN":
                    filtered_jpn = [r for r in filtered_jpn if str(r.get('peperiksaan','')).strip().upper() == str(selected_pep_filter).strip().upper()]
                
                if not filtered_jpn:
                    if selected_ppd_filter != "SEMUA DAERAH" or selected_pep_filter != "SEMUA PEPERIKSAAN":
                        st.warning(f"📭 Belum ada laporan untuk Daerah={selected_ppd_filter}, Tarikh={selected_tarikh_filter}, Peperiksaan={selected_pep_filter}")
                        ppd_counts_real = Counter([KOD_PPD_MAPPING.get(str(r.get('kod_ppd','') or r.get('no_pusat','')[:2]).upper(), r.get('ppd','')) for r in reports_for_download])
                        pep_counts_real = Counter([str(r.get('peperiksaan','')).strip().upper() for r in reports_for_download])
                        st.info(f"Filter tiada data. Total sistem {len(reports_for_download)} laporan: PPD={', '.join([f'{k}={v}' for k,v in ppd_counts_real.items()])} | Pep={', '.join([f'{k}={v}' for k,v in pep_counts_real.items()])}")
                    else:
                        st.warning(f"Tiada laporan untuk Daerah={selected_ppd_filter}, Tarikh={selected_tarikh_filter}, Peperiksaan={selected_pep_filter}")
                else:
                    st.success(f"✅ {len(filtered_jpn)} laporan untuk Daerah={selected_ppd_filter}, Tarikh={selected_tarikh_filter}, Peperiksaan={selected_pep_filter} akan dijana")
                    with st.expander(f"📋 Lihat {len(filtered_jpn)} laporan", expanded=False):
                        for r in filtered_jpn[:20]:
                            st.write(f"- {r.get('tarikh')} | {r.get('no_pusat')} | {r.get('nama_pusat','')[:30]} | {r.get('ppd')} | {r.get('peperiksaan','')} | {r.get('nama_kp','')[:20]}")
                    
                    col_jana, col_dl = st.columns(2)
                    with col_jana:
                        if st.button(f"🔄 Jana ZIP {len(filtered_jpn)} Laporan", key="jpn_jana_filtered_v2", type="primary", use_container_width=True):
                            with st.spinner(f"Jana {len(filtered_jpn)} PDF..."):
                                try:
                                    import zipfile
                                    from io import BytesIO
                                    zip_buffer = BytesIO()
                                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                        for r in filtered_jpn:
                                            no_pusat = r.get('no_pusat','UNKNOWN')
                                            safe_nama = "".join(c if c.isalnum() or c in ('-','_') else '_' for c in r.get('nama_pusat','')[:15])
                                            base_filename = f"{r.get('tarikh','')}_{r.get('ppd','')}_{no_pusat}_{safe_nama}"
                                            try:
                                                pdf_data = jana_blpp_pdf(r)
                                                if isinstance(pdf_data, BytesIO):
                                                    pdf_bytes = pdf_data.getvalue()
                                                else:
                                                    pdf_bytes = pdf_data
                                                zip_file.writestr(f"{base_filename}.pdf", pdf_bytes)
                                            except:
                                                try:
                                                    html_content = jana_blpp_html(r)
                                                    zip_file.writestr(f"{base_filename}.html", html_content)
                                                except:
                                                    pass
                                    zip_buffer.seek(0)
                                    st.session_state["jpn_filtered_zip"] = zip_buffer.getvalue()
                                    st.session_state["jpn_filtered_count"] = len(filtered_jpn)
                                    st.success(f"✅ {len(filtered_jpn)} PDF dijana!")
                                except Exception as e:
                                    st.error(f"Error: {e}")
                    with col_dl:
                        if "jpn_filtered_zip" in st.session_state:
                            count = st.session_state.get("jpn_filtered_count", len(filtered_jpn))
                            safe_pep = "".join(c if c.isalnum() else "_" for c in str(selected_pep_filter)[:20])
                            st.download_button(
                                label=f"⬇️ Download ZIP ({count}) - {selected_pep_filter}",
                                data=st.session_state["jpn_filtered_zip"],
                                file_name=f"BLPP_JPN_{selected_ppd_filter}_{selected_tarikh_filter}_{safe_pep}_{count}laporan.zip",
                                mime="application/zip",
                                key="jpn_dl_filtered_v2",
                                type="primary",
                                use_container_width=True
                            )
            
            # PPD: Pilihan Tarikh & Peperiksaan (v38 - KASHAH: PPD pun sama, pilih peperiksaan & tarikh, tak campur)
            if peranan_user == "PPD":
                st.markdown("---")
                st.markdown(f"#### 🎯 Pilih Tarikh & Peperiksaan (PPD {user.get('ppd')} - Daerah Anda Sahaja)")
                col_tarikh_ppd, col_pep_ppd = st.columns(2)
                with col_tarikh_ppd:
                    all_tarikhs_ppd = sorted(list(set([r.get('tarikh','') for r in reports_for_download if r.get('tarikh')])), reverse=True)
                    tarikh_options_ppd = ["SEMUA TARIKH"] + all_tarikhs_ppd
                    selected_tarikh_ppd = st.selectbox("Pilih Tarikh", tarikh_options_ppd, key="ppd_tarikh_filter_v38")
                with col_pep_ppd:
                    all_peps_ppd = sorted(list(set([str(r.get('peperiksaan','')).strip().upper() for r in reports_for_download if r.get('peperiksaan')])))
                    pep_options_ppd = ["SEMUA PEPERIKSAAN"] + (PEPERIKSAAN_LIST if 'PEPERIKSAAN_LIST' in globals() else all_peps_ppd)
                    seen_ppd = set()
                    pep_unique_ppd = []
                    for p in pep_options_ppd:
                        pu = str(p).strip().upper()
                        if pu not in seen_ppd:
                            seen_ppd.add(pu)
                            pep_unique_ppd.append(p)
                    for ap in all_peps_ppd:
                        if ap not in seen_ppd:
                            pep_unique_ppd.append(ap)
                    selected_pep_ppd = st.selectbox("Pilih Peperiksaan", pep_unique_ppd, key="ppd_pep_filter_v38")
                
                filtered_ppd = reports_for_download
                if selected_tarikh_ppd != "SEMUA TARIKH":
                    filtered_ppd = [r for r in filtered_ppd if r.get('tarikh') == selected_tarikh_ppd]
                if selected_pep_ppd != "SEMUA PEPERIKSAAN":
                    filtered_ppd = [r for r in filtered_ppd if str(r.get('peperiksaan','')).strip().upper() == str(selected_pep_ppd).strip().upper()]
                
                if not filtered_ppd:
                    st.warning(f"📭 Tiada laporan untuk Tarikh={selected_tarikh_ppd}, Peperiksaan={selected_pep_ppd} di daerah {user.get('ppd')}")
                else:
                    st.success(f"✅ {len(filtered_ppd)} laporan untuk Daerah={user.get('ppd')}, Tarikh={selected_tarikh_ppd}, Peperiksaan={selected_pep_ppd} akan dijana")
                    with st.expander(f"📋 Lihat {len(filtered_ppd)} laporan", expanded=False):
                        for r in filtered_ppd[:20]:
                            st.write(f"- {r.get('tarikh')} | {r.get('no_pusat')} | {r.get('nama_pusat','')[:30]} | {r.get('peperiksaan','')} | {r.get('nama_kp','')[:20]}")
                
                    col_jana_ppd, col_dl_ppd = st.columns(2)
                    with col_jana_ppd:
                        if st.button(f"🔄 Jana ZIP {len(filtered_ppd)} Laporan", key="ppd_jana_filtered_v38", type="primary", use_container_width=True):
                            with st.spinner(f"Jana {len(filtered_ppd)} PDF..."):
                                try:
                                    import zipfile
                                    from io import BytesIO
                                    zip_buffer = BytesIO()
                                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                        for r in filtered_ppd:
                                            no_pusat = r.get('no_pusat','UNKNOWN')
                                            safe_nama = "".join(c if c.isalnum() or c in ('-','_') else '_' for c in r.get('nama_pusat','')[:15])
                                            base_filename = f"{r.get('tarikh','')}_{r.get('ppd','')}_{no_pusat}_{safe_nama}"
                                            try:
                                                pdf_data = jana_blpp_pdf(r)
                                                if isinstance(pdf_data, BytesIO):
                                                    pdf_bytes = pdf_data.getvalue()
                                                else:
                                                    pdf_bytes = pdf_data
                                                zip_file.writestr(f"{base_filename}.pdf", pdf_bytes)
                                            except:
                                                try:
                                                    html_content = jana_blpp_html(r)
                                                    zip_file.writestr(f"{base_filename}.html", html_content)
                                                except:
                                                    pass
                                    zip_buffer.seek(0)
                                    st.session_state["ppd_filtered_zip"] = zip_buffer.getvalue()
                                    st.session_state["ppd_filtered_count"] = len(filtered_ppd)
                                    st.success(f"✅ {len(filtered_ppd)} PDF dijana!")
                                except Exception as e:
                                    st.error(f"Error: {e}")
                    with col_dl_ppd:
                        if "ppd_filtered_zip" in st.session_state:
                            count = st.session_state.get("ppd_filtered_count", len(filtered_ppd))
                            safe_pep_ppd = "".join(c if c.isalnum() else "_" for c in str(selected_pep_ppd)[:20])
                            st.download_button(
                                label=f"⬇️ Download ZIP ({count}) - {selected_pep_ppd}",
                                data=st.session_state["ppd_filtered_zip"],
                                file_name=f"BLPP_PPD_{user.get('ppd')}_{selected_tarikh_ppd}_{safe_pep_ppd}_{count}laporan.zip",
                                mime="application/zip",
                                key="ppd_dl_filtered_v38",
                                type="primary",
                                use_container_width=True
                            )
            
            st.markdown("---")
            st.markdown("#### 📅 Pilih Tarikh untuk Download Pukal (Ikut Tarikh)")
            for t, c in sorted(tarikh_counts.items(), reverse=True):
                with st.container(border=True):
                    col1, col2, col3 = st.columns([2,1,1])
                    with col1:
                        st.write(f"📅 **{t}** - {c} laporan")
                        from collections import Counter
                        def get_ppd_betul(r):
                            kod = r.get('kod_ppd','') or str(r.get('no_pusat',''))[:2].upper()
                            if kod in KOD_PPD_MAPPING:
                                return KOD_PPD_MAPPING[kod]
                            return r.get('ppd','') or 'TIADA PPD'
                        ppd_counter = Counter([get_ppd_betul(r) for r in reports_for_download if r.get('tarikh')==t])
                        ppd_detail = ", ".join([f"{ppd}({count})" for ppd, count in ppd_counter.items()])
                        st.caption(f"PPD: {ppd_detail}")
                    with col2:
                        if st.button(f"🔄 Jana ZIP", key=f"bulk_menu_jana_{t}_v2", type="primary", use_container_width=True):
                            with st.spinner(f"Jana {c} PDF..."):
                                try:
                                    zip_data, count = create_bulk_pdf_zip_by_date(reports_for_download, t)
                                    if zip_data:
                                        st.session_state[f"bulk_menu_zip_{t}_v2"] = zip_data
                                        st.session_state[f"bulk_menu_count_{t}_v2"] = count
                                        st.success(f"✅ {count} PDF dijana!")
                                    else:
                                        st.error("Gagal jana ZIP")
                                except Exception as e:
                                    st.error(f"Error: {e}")
                    with col3:
                        zk = f"bulk_menu_zip_{t}_v2"
                        if zk in st.session_state:
                            count = st.session_state.get(f"bulk_menu_count_{t}_v2", c)
                            st.download_button(
                                label=f"⬇️ Download ZIP ({count})", 
                                data=st.session_state[zk], 
                                file_name=f"BLPP_PUKAL_{t}_{count}laporan.zip", 
                                mime="application/zip", 
                                key=f"bulk_menu_dl_{t}_v2",
                                use_container_width=True
                            )
                        else:
                            st.caption("Tekan Jana ZIP dulu")
            
            st.markdown("---")
            st.markdown("#### 📥 Download Semua Tarikh (All-in-One)")
            if st.button("🔄 Jana ZIP Semua Laporan", key="bulk_all_jana_v2", type="secondary"):
                with st.spinner(f"Jana {len(reports_for_download)} PDF..."):
                    try:
                        import zipfile
                        from io import BytesIO
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for r in reports_for_download:
                                no_pusat = r.get('no_pusat','UNKNOWN')
                                safe_nama = "".join(c if c.isalnum() or c in ('-','_') else '_' for c in r.get('nama_pusat','')[:15])
                                base_filename = f"{r.get('tarikh','')}_{r.get('ppd','')}_{no_pusat}_{safe_nama}"
                                try:
                                    pdf_data = jana_blpp_pdf(r)
                                    if isinstance(pdf_data, BytesIO):
                                        pdf_bytes = pdf_data.getvalue()
                                    else:
                                        pdf_bytes = pdf_data
                                    zip_file.writestr(f"{base_filename}.pdf", pdf_bytes)
                                except:
                                    try:
                                        html_content = jana_blpp_html(r)
                                        zip_file.writestr(f"{base_filename}.html", html_content)
                                    except:
                                        pass
                        zip_buffer.seek(0)
                        st.session_state["bulk_all_zip_v2"] = zip_buffer.getvalue()
                        st.success(f"✅ {len(reports_for_download)} PDF dijana!")
                    except Exception as e:
                        st.error(f"Error: {e}")
            
            if "bulk_all_zip_v2" in st.session_state:
                st.download_button(
                    label=f"⬇️ Download ZIP Semua ({len(reports_for_download)} laporan)",
                    data=st.session_state["bulk_all_zip_v2"],
                    file_name=f"BLPP_PUKAL_SEMUA_{len(reports_for_download)}laporan.zip",
                    mime="application/zip",
                    key="bulk_all_dl_v2"
                )

    # ========== MENU: PADAM LAPORAN
    # ========== MENU: PADAM LAPORAN - JPN SAHAJA ==========
    elif menu.startswith("🗑️"):

        st.markdown('<div class="card"><h3>🗑️ Padam Laporan - JPN Selangor</h3><p style="font-size:12px;">Function delete: satu persatu, ikut hari, ikut daerah, keseluruhan</p></div>', unsafe_allow_html=True)
        
        if peranan_user != "ADMIN_JPN":
            st.error("Hanya JPN boleh padam laporan!")
        else:
            if not reports:
                st.info("Tiada laporan untuk dipadam")
            else:
                tab1, tab2, tab3, tab4 = st.tabs(["🗑️ Satu Persatu", "📅 Ikut Hari", "📍 Ikut Daerah", "💣 Keseluruhan"])
                
                with tab1:
                    st.markdown("### Padam Satu Persatu")
                    st.caption(f"Total {len(reports)} laporan")
                    report_options_del = {}
                    for r in sorted(reports, key=lambda x: x.get('tarikh_submit',''), reverse=True):
                        sidang_l = r.get('sidang','')
                        label = f"{r.get('tarikh','')} | {sidang_l} | {r.get('no_pusat','')} | {r.get('nama_pusat','')} | {r.get('ppd','')} | {r.get('id','')}"
                        report_options_del[label] = r
                    
                    sel_del = st.selectbox("Pilih Laporan untuk Dipadam", list(report_options_del.keys()), key="del_one")
                    if sel_del:
                        r_del = report_options_del[sel_del]
                        st.warning(f"Anda akan padam: {sel_del}")
                        st.json({k: r_del.get(k,'') for k in ['id','tarikh','sidang','no_pusat','nama_pusat','ppd','tarikh_submit']})
                        confirm = st.checkbox("Saya confirm nak padam laporan ini", key="confirm_one")
                        if st.button("🗑️ Padam Laporan Ini", type="primary", key="btn_del_one"):
                            if confirm:
                                reports = [r for r in reports if r.get('id') != r_del.get('id')]
                                save_reports(reports)
                                st.success(f"✅ Laporan {r_del.get('id')} telah dipadam!")
                                st.rerun()
                            else:
                                st.error("Sila tick confirm dulu")
                
                with tab2:
                    st.markdown("### Padam Ikut Hari")
                    tarikh_list = sorted(list(set([r.get('tarikh','') for r in reports if r.get('tarikh')])), reverse=True)
                    if tarikh_list:
                        sel_tarikh_del = st.selectbox("Pilih Tarikh", tarikh_list, key="del_tarikh")
                        count_tarikh = len([r for r in reports if r.get('tarikh')==sel_tarikh_del])
                        st.info(f"📅 {count_tarikh} laporan pada {sel_tarikh_del} akan dipadam")
                        # Preview
                        preview = [r for r in reports if r.get('tarikh')==sel_tarikh_del][:5]
                        for r in preview:
                            st.caption(f"- {r.get('no_pusat')} | {r.get('nama_pusat')} | {r.get('sidang','')} | {r.get('id')}")
                        if count_tarikh > 5:
                            st.caption(f"... dan {count_tarikh-5} lagi")
                        
                        confirm2 = st.checkbox(f"Saya confirm nak padam SEMUA {count_tarikh} laporan pada {sel_tarikh_del}", key="confirm_tarikh")
                        if st.button(f"🗑️ Padam Semua Laporan Tarikh {sel_tarikh_del}", type="primary", key="btn_del_tarikh"):
                            if confirm2:
                                reports = [r for r in reports if r.get('tarikh') != sel_tarikh_del]
                                save_reports(reports)
                                st.success(f"✅ {count_tarikh} laporan pada {sel_tarikh_del} telah dipadam!")
                                st.rerun()
                            else:
                                st.error("Sila tick confirm dulu")
                    else:
                        st.info("Tiada tarikh")
                
                with tab3:
                    st.markdown("### Padam Ikut Daerah (PPD)")
                    ppd_list_del = sorted(list(set([r.get('ppd','') for r in reports if r.get('ppd')])))
                    if ppd_list_del:
                        sel_ppd_del = st.selectbox("Pilih PPD/Daerah", ppd_list_del, key="del_ppd")
                        count_ppd = len([r for r in reports if r.get('ppd')==sel_ppd_del])
                        st.info(f"📍 {count_ppd} laporan PPD {sel_ppd_del} akan dipadam")
                        preview_ppd = [r for r in reports if r.get('ppd')==sel_ppd_del][:5]
                        for r in preview_ppd:
                            st.caption(f"- {r.get('tarikh')} | {r.get('no_pusat')} | {r.get('nama_pusat')} | {r.get('sidang','')} | {r.get('id')}")
                        if count_ppd > 5:
                            st.caption(f"... dan {count_ppd-5} lagi")
                        
                        confirm3 = st.checkbox(f"Saya confirm nak padam SEMUA {count_ppd} laporan PPD {sel_ppd_del}", key="confirm_ppd")
                        if st.button(f"🗑️ Padam Semua Laporan PPD {sel_ppd_del}", type="primary", key="btn_del_ppd"):
                            if confirm3:
                                reports = [r for r in reports if r.get('ppd') != sel_ppd_del]
                                save_reports(reports)
                                st.success(f"✅ {count_ppd} laporan PPD {sel_ppd_del} telah dipadam!")
                                st.rerun()
                            else:
                                st.error("Sila tick confirm dulu")
                    else:
                        st.info("Tiada PPD")
                
                with tab4:
                    st.markdown("### Padam Keseluruhan Laporan")
                    st.error(f"⚠️ AMARAN: Ini akan padam SEMUA {len(reports)} laporan dalam sistem! Tidak boleh undo!")
                    st.json({
                        "total_laporan": len(reports),
                        "tarikh_ada": sorted(list(set([r.get('tarikh','') for r in reports]))),
                        "ppd_ada": sorted(list(set([r.get('ppd','') for r in reports])))
                    })
                    confirm_text = st.text_input("Taip 'PADAM SEMUA' untuk confirm", key="confirm_all_text")
                    confirm4 = st.checkbox(f"Saya FAHAM dan confirm nak padam SEMUA {len(reports)} laporan", key="confirm_all")
                    if st.button("💣 PADAM KESELURUHAN LAPORAN", type="primary", key="btn_del_all"):
                        if confirm_text == "PADAM SEMUA" and confirm4:
                            total = len(reports)
                            reports = []
                            save_reports(reports)
                            st.success(f"✅ {total} laporan telah dipadam keseluruhan!")
                            st.rerun()
                        else:
                            st.error("Sila taip 'PADAM SEMUA' dan tick confirm")
    

    # ========== MENU 4: PENGURUSAN PENGGUNA (JPN) / STATISTIK ==========
    else:
        if peranan_user == "ADMIN_JPN":
            st.markdown('<div class="card"><h3>👥 Pengurusan Pengguna - JPN Selangor (Admin)</h3><p style="font-size:12px;">Admin JPN bertindak sebagai pengelola, menjaga sistem, selenggara sistem. 504 KP + 10 PPD + LP.</p></div>', unsafe_allow_html=True)
        
            tab1, tab2, tab3 = st.tabs(["📊 Dashboard JPN", "👥 Senarai 504 KP", "➕ Tambah KP / Import"])
        
            with tab1:
                col1, col2, col3, col4 = st.columns(4)
                total_kp = len([u for u in load_users() if u["peranan"]=="KP"])
                total_laporan = len(reports)
                hari_ini = len([r for r in reports if r.get("tarikh")==date.today().strftime("%Y-%m-%d")])
                ada_kes = len([r for r in reports if not r.get("tiada_kes")])
                with col1:
                    st.metric("Total KP", f"{total_kp} / 504")
                with col2:
                    st.metric("Total Laporan", total_laporan)
                with col3:
                    st.metric("Laporan Hari Ini", hari_ini)
                with col4:
                    st.metric("Ada Kes", ada_kes)
            
                st.markdown("#### Laporan Ikut PPD (10 PPD)")
                ppd_stats = []
                for ppd in PPD_LIST:
                    # v31 JPN STRICT: guna helper yang sama macam PPD - strict by No Pusat prefix
                    count = len([r for r in reports if is_report_for_ppd_strict(r, ppd)])
                    ppd_stats.append({"PPD": ppd, "Laporan": count, "KP": len([u for u in load_users() if u.get("ppd")==ppd and u["peranan"]=="KP"])})
                st.dataframe(pd.DataFrame(ppd_stats), use_container_width=True)
            
                # Chart
                if reports:
                    df_chart = pd.DataFrame(reports)
                    st.bar_chart(df_chart.groupby("ppd").size())
        
            with tab2:
                users = load_users()
                df_users = pd.DataFrame([u for u in users if u["peranan"]=="KP"])
                if not df_users.empty:
                    st.dataframe(df_users, use_container_width=True, height=500)
                    # Download
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_users.to_excel(writer, index=False)
                    st.download_button("📥 Download Senarai KP Excel", output.getvalue(), file_name="Senarai_504_KP.xlsx")
                else:
                    st.info("Tiada KP")
        
            with tab3:
                st.markdown("#### Tambah KP Baru")
                with st.form("add_kp"):
                    c1, c2 = st.columns(2)
                    with c1:
                        new_id = st.text_input("ID KP (KP001-504)")
                        new_nama = st.text_input("Nama KP")
                        new_ppd = st.selectbox("PPD", PPD_LIST)
                    with c2:
                        new_no_pusat = st.text_input("No. Pusat")
                        new_nama_pusat = st.text_input("Nama Pusat")
                        new_pass = st.text_input("Password", value="kp123")
                    if st.form_submit_button("Tambah KP"):
                        users = load_users()
                        users.append({"id":new_id,"nama":new_nama,"peranan":"KP","ppd":new_ppd,"no_pusat":new_no_pusat,"nama_pusat":new_nama_pusat,"password":new_pass})
                        save_users(users)
                        st.success(f"KP {new_id} ditambah!")
            
                st.markdown("---")
                st.markdown("#### 📤 Import Bulk 504 KP dari Excel/CSV")
                st.info("Format CSV: id,nama,peranan,ppd,no_pusat,nama_pusat,password")
                uploaded = st.file_uploader("Upload CSV", type=["csv"])
                if uploaded:
                    df_import = pd.read_csv(uploaded)
                    st.dataframe(df_import.head())
                    if st.button("Import Sekarang"):
                        users = load_users()
                        for _, row in df_import.iterrows():
                            users.append(row.to_dict())
                        save_users(users)
                        st.success(f"{len(df_import)} KP diimport!")
    
        else:
            # Statistik untuk PPD / LP / KP - FIX robust Petaling Utama
            st.markdown(f'<div class="card"><h3>📈 Statistik {peranan_user} - {user.get("ppd","")}</h3></div>', unsafe_allow_html=True)
        
            my_reports = reports
            if peranan_user == "PPD":
                user_ppd = user["ppd"]
                allowed_kods = get_kods_for_ppd(user_ppd)
                # v29 STRICT STATISTIK: PPD Klang hanya kira BA, bukan BB/BC etc - ikut No Pusat prefix
                my_reports = [r for r in reports if is_report_for_ppd_strict(r, user_ppd)]
                if len(my_reports)==0 and len(reports)>0:
                    st.info(f"Belum ada laporan untuk PPD {user_ppd}. KP daerah {user_ppd} belum hantar lagi.")
            elif peranan_user == "KP":
                my_reports = [r for r in reports if r.get("id_kp")==user["id"]]
        
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Laporan", len(my_reports))
            with col2:
                st.metric("Tiada Kes", len([r for r in my_reports if r.get("tiada_kes")]))
            with col3:
                st.metric("Ada Kes", len([r for r in my_reports if not r.get("tiada_kes")]))
        
            if my_reports:
                df = pd.DataFrame(my_reports)
                st.markdown("#### Trend Laporan Harian (20 Hari Peperiksaan)")
                daily = df.groupby("tarikh").size()
                st.bar_chart(daily)
            
                st.markdown("#### Laporan Terkini")
                st.dataframe(df.sort_values("tarikh", ascending=False).head(10), use_container_width=True)

    # Footer - REMOVED FOR CLEAN PRINT - No watermark
    # Borang rasmi 100% bersih tanpa nama kecik
    st.markdown('<style>@media print { footer, [data-testid="stFooter"], .footer-global { display:none !important; } }</style>', unsafe_allow_html=True)
