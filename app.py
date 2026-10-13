import streamlit as st
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

def save_to_gsheets(report):
    client = get_gsheets_client()
    if not client:
        return False
    try:
        # ID Sheet dari secrets atau default
        sheet_id = st.secrets.get("gsheets_id", "") if "gsheets_id" in st.secrets else ""
        if not sheet_id:
            return False
        sh = client.open_by_key(sheet_id)
        ws = sh.worksheet("BLPP_Reports")
        # Convert kes_data to JSON string
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
            report.get("tandatangan_b64","")[:5000],  # Simpan sebahagian saja untuk demo, full di Drive
            report.get("telefon_kp","")
        ]
        ws.append_row(row)
        return True
    except Exception as e:
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


def get_all_pusat():
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

PPD_TO_KOD = {v:k for k,v in KOD_PPD_MAPPING.items()}

PEPERIKSAAN_LIST = ["SPM 2026", "SPM Ulangan 2026", "STPM 2026", "MUET 2026", "Ujian Amali Sains SPM", "STAM 2026"]

JENIS_PEPERIKSAAN = ["Peperiksaan SPM (504 Pusat)", "Ujian Amali Sains (277 Pusat Makmal)"]

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
    # 2. Cari Excel SPM - fleksibel nama ada (1) (3) pun boleh
    search_folders = ["data", "/tmp/blpp_data", ".", "./data"]
    for folder in search_folders:
        if not os.path.exists(folder):
            continue
        # Cari file yang ada v1.1 atau spm
        for file in os.listdir(folder):
            if file.endswith('.xlsx') and ('v1.1' in file or 'spm' in file.lower() or 'pusat' in file.lower() and 'amali' not in file.lower()):
                try:
                    path = os.path.join(folder, file)
                    df = pd.read_excel(path, sheet_name="selenggara_pusat")
                    centers = []
                    for _, row in df.iterrows():
                        kod = str(row["Kod_PPD"]).strip()
                        centers.append({
                            "Kod_PPD": kod,
                            "No_Pusat": str(row["No_Pusat"]).strip(),
                            "Nama_Pusat": str(row["Nama_Pusat"]).strip(),
                            "PPD": KOD_PPD_MAPPING.get(kod, kod),
                            "Bil_Calon": int(row["Bil_Calon_Pusat"]) if pd.notna(row["Bil_Calon_Pusat"]) else 0
                        })
                    if centers:
                        return centers
                except Exception as e:
                    continue
        # Try any xlsx with selenggara_pusat sheet
        for file in os.listdir(folder):
            if file.endswith('.xlsx'):
                try:
                    path = os.path.join(folder, file)
                    xls = pd.ExcelFile(path)
                    if "selenggara_pusat" in xls.sheet_names:
                        df = pd.read_excel(path, sheet_name="selenggara_pusat")
                        centers = []
                        for _, row in df.iterrows():
                            kod = str(row["Kod_PPD"]).strip()
                            centers.append({
                                "Kod_PPD": kod,
                                "No_Pusat": str(row["No_Pusat"]).strip(),
                                "Nama_Pusat": str(row["Nama_Pusat"]).strip(),
                                "PPD": KOD_PPD_MAPPING.get(kod, kod),
                                "Bil_Calon": int(row["Bil_Calon_Pusat"]) if pd.notna(row["Bil_Calon_Pusat"]) else 0
                            })
                        if centers:
                            return centers
                except:
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
REPORTS_FILE = os.path.join(DATA_DIR, "blpp_reports.json")

def load_users():
    # KP001 style - 504 KP full - FIX: generate 504 bukan 20 saja
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                data = json.load(f)
                # kalau dah ada file tapi masih format lama 20 saja, regenerate ke 504
                kp_count = len([u for u in data if u.get("peranan")=="KP"])
                if kp_count >= 500:
                    return data
                else:
                    os.remove(USERS_FILE)
        except:
            pass

    # Default users - 504 KP + admin
    users = []
    # Admin JPN
    users.append({"id":"JPN001","nama":"Admin JPN Selangor","peranan":"ADMIN_JPN","ppd":"JPN","no_pusat":"-","password":"jpn123","telefon":""})
    users.append({"id":"JPN002","nama":"JPN Selangor 2","peranan":"ADMIN_JPN","ppd":"JPN","no_pusat":"-","password":"jpn123","telefon":""})
    # LP
    users.append({"id":"LP001","nama":"Lembaga Peperiksaan","peranan":"LP","ppd":"Semua","no_pusat":"-","password":"lp123","telefon":""})
    # 10 PPD
    for i, ppd in enumerate(PPD_LIST):
        users.append({"id":f"PPD{i+1:02d}","nama":f"PPD {ppd}","peranan":"PPD","ppd":ppd,"no_pusat":"-","password":f"ppd{i+1:02d}","telefon":""})

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
            uid = f"KP{idx:03d}"
            users.append({
                "id": uid,
                "nama": f"KP {uid} - {nama_pusat[:30]}",
                "peranan": "KP",
                "ppd": ppd,
                "no_pusat": no_pusat,
                "nama_pusat": nama_pusat,
                "password": uid.lower(),
                "telefon": ""
            })
    else:
        # Fallback 504 ikut mapping BA-BK
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
            uid = f"KP{i:03d}"
            users.append({
                "id": uid,
                "nama": f"KP {uid} - {nama_pusat[:30]}",
                "peranan": "KP",
                "ppd": ppd,
                "no_pusat": no_pusat,
                "nama_pusat": nama_pusat,
                "password": uid.lower(),
                "telefon": ""
            })

    # Legacy BH562 / BC001 tetap boleh login
    users.append({"id":"BH562","nama":"MAAHAD SAINS AL-QURAN AL-AMIN (Legacy)","peranan":"KP","ppd":"Petaling Perdana","no_pusat":"BH562","nama_pusat":"MAAHAD SAINS AL-QURAN AL-AMIN","password":"bh562","telefon":""})
    users.append({"id":"BC001","nama":"SMK Sepang (Legacy)","peranan":"KP","ppd":"Sepang","no_pusat":"BC001","nama_pusat":"SMK Sepang","password":"bc001","telefon":""})
    return users


def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)

def load_reports():
    if os.path.exists(REPORTS_FILE):
        with open(REPORTS_FILE, "r") as f:
            return json.load(f)
    return []

def save_reports(reports):
    with open(REPORTS_FILE, "w") as f:
        json.dump(reports, f, indent=2)

# ========== FUNGSI JANA BORANG BLPP HTML (OFFICIAL) ==========




def jana_blpp_html(report):
    no_pusat = report.get('no_pusat','')
    nama_pusat = report.get('nama_pusat','')
    nama_kp = report.get('nama_kp','')
    peperiksaan = report.get('peperiksaan','')
    tahun = report.get('tahun','')
    kes_data = report.get('kes_data', [])
    tiada_kes = report.get('tiada_kes', False)
    
    if peperiksaan and tahun:
        pep_line = f"PEPERIKSAAN {peperiksaan} TAHUN {tahun}"
    else:
        pep_line = "PEPERIKSAAN ........................................ TAHUN ........................................"
    
    if tiada_kes or not kes_data:
        kes_rows = ""
        for _ in range(6):
            kes_rows += "<tr><td style='height:32px;'>&nbsp;</td><td></td><td></td><td></td><td></td></tr>"
    else:
        kes_rows = ""
        for i, k in enumerate(kes_data, 1):
            kes_rows += f"<tr><td>{i}</td><td>{k.get('angka_giliran','')}</td><td>{k.get('no_ic','')}</td><td style='text-align:left;'>{k.get('kes','')}</td><td style='text-align:left;'>{k.get('tindakan','')}</td></tr>"
        for _ in range(max(0, 6 - len(kes_data))):
            kes_rows += "<tr><td style='height:32px;'>&nbsp;</td><td></td><td></td><td></td><td></td></tr>"

    html = f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
        @page {{ size: A4 landscape; margin: 10mm 12mm 8mm 12mm; }}
        body {{ font-family: Arial, sans-serif; margin:0; }}
        .page {{ width:100%; max-width:1020px; margin:0 auto; }}
        .hdr {{ display:flex; justify-content:space-between; }}
        .hdr-center {{ text-align:center; font-weight:bold; font-size:10.5pt; flex:1; line-height:1.2; }}
        .hdr-right {{ text-align:right; font-size:8.5pt; font-weight:bold; }}
        hr {{ border:none; border-top:1.5px solid #000; margin:6px 0; }}
        .pep {{ text-align:center; font-weight:bold; font-size:11pt; margin:8px 0 2px 0; }}
        .title {{ text-align:center; font-weight:bold; font-size:12.5pt; text-decoration:underline; margin:2px 0 8px 0; }}
        /* FIX: arahan sama lebar dengan table */
        .arahan {{ width:100%; font-size:7.8pt; text-align:justify; line-height:1.2; box-sizing:border-box; }}
        .arahan ol {{ margin:0; padding-left:16px; }}
        table {{ width:100%; border-collapse:collapse; }}
        th, td {{ border:1px solid #000; padding:3px 5px; font-size:8.5pt; }}
        th {{ text-align:center; font-weight:bold; }}
        td {{ text-align:center; }}
        .footer {{ display:flex; justify-content:space-between; margin-top:20px; font-size:8.5pt; }}
    </style></head><body>
    <div class="page">
        <div class="hdr">
            <div style="width:90px;"></div>
            <div class="hdr-center">LEMBAGA PEPERIKSAAN<br>KEMENTERIAN PENDIDIKAN MALAYSIA</div>
            <div class="hdr-right">LP/Am 90 Pin. 1/2023<br><b style="font-size:12pt;">BLPP</b></div>
        </div>
        <hr>
        <div class="pep">{pep_line}</div>
        <div class="title">BORANG LAPORAN PENGENDALIAN PEPERIKSAAN</div>
        <div class="arahan"><ol>
            <li>Borang ini hendaklah digunakan semasa peperiksaan untuk melaporkan mengenai perkara-perkara di luar dugaan yang berlaku kepada calon seperti sakit secara tiba-tiba, kemalangan, ditahan polis dan lain-lain.</li>
            <li>Jika ada kes, Ketua Pengawas Peperiksaan/Ketua Pengawas Amali/Ketua Pentaksir/Assessor dikehendaki mengisi borang ini dan menyerahkannya kepada Penyelia Kawasan untuk dihantar melalui e-mel pada hari yang sama kepada Pengarah Peperiksaan, Lembaga Peperiksaan (u.p.: Unit Menengah Atas/Unit Pra Universiti/Unit Vokasional) dan Pengarah Pendidikan Negeri, Jabatan Pendidikan Negeri (u.p.: Sektor Pentaksiran dan Peperiksaan).</li>
        </ol></div>
        <table style="margin-top:6px;">
            <tr><th style="width:14%;">NO. PUSAT</th><th style="width:34%;">NAMA PUSAT PEPERIKSAAN</th><th style="width:52%;">NAMA KETUA PENGAWAS PEPERIKSAAN/KETUA PENGAWAS AMALI/KETUA PENTAKSIR/ASSESSOR</th></tr>
            <tr><td style="font-weight:bold; height:22px;">{no_pusat}</td><td style="font-weight:bold;">{nama_pusat}</td><td style="font-weight:bold;">{nama_kp}</td></tr>
        </table>
        <table style="margin-top:8px;">
            <tr><th style="width:5%;">BIL.</th><th style="width:15%;">ANGKA GILIRAN</th><th style="width:18%;">NO. PENGENALAN DIRI</th><th style="width:31%;">KES</th><th style="width:31%;">TINDAKAN</th></tr>
            {kes_rows}
        </table>
        <div class="footer">
            <div>Pelapor:<br><br><br>___________________________________________<br>Tandatangan Ketua Pengawas Peperiksaan/<br>Ketua Pengawas Amali/Ketua Pentaksir/Assessor<br><br>Tarikh: __________________________________________</div>
            <div>Pengesahan:<br><br><br>______________________________________________<br>Tandatangan Penyelia Kawasan/Pegawai Peperiksaan<br><br><br>Tarikh: ____________________________________________</div>
        </div>
    </div></body></html>
    """
    return html




def jana_blpp_pdf(report):
    """PDF SEBIJIK 100% BORANG ASAL LP - FINAL - logo besar, tanpa garisan, tengah 1 kotak besar kosong"""
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
        for lp in ["logo_lp.png", "data/logo_lp.png", "./logo_lp.png", "/tmp/logo_final.png"]:
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

        if logo_to_use:
            try:
                pdf.image(logo_to_use, x=10, y=4, w=32)
            except:
                pass

        pdf.set_font("Helvetica", "B", 10)
        pdf.set_xy(45, 7)
        pdf.cell(190, 5, "LEMBAGA PEPERIKSAAN", align="C")
        pdf.set_xy(45, 12)
        pdf.cell(190, 5, "KEMENTERIAN PENDIDIKAN MALAYSIA", align="C")
        pdf.set_font("Helvetica", "", 7)
        pdf.set_xy(250, 7)
        pdf.cell(37, 4, "LP/Am 90 Pin. 1/2023", align="R")
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_xy(250, 11)
        pdf.cell(37, 7, "BLPP", align="R")

        pep_text = f"PEPERIKSAAN {pep} TAHUN {thn}" if pep and thn else "PEPERIKSAAN ................................... TAHUN ....................."
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_xy(10, 22)
        pdf.cell(277, 5, pep_text, align="C")
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_xy(10, 28)
        pdf.cell(277, 6, "BORANG LAPORAN PENGENDALIAN PEPERIKSAAN", align="C")

        pdf.set_font("Helvetica", "", 6.5)
        pdf.set_xy(10, 36)
        arahan = """1. Borang ini hendaklah digunakan semasa peperiksaan untuk melaporkan mengenai perkara-perkara di luar dugaan yang berlaku kepada calon seperti sakit secara tiba-tiba, kemalangan, ditahan polis dan lain-lain.
2. Jika ada kes, Ketua Pengawas Peperiksaan/Ketua Pengawas Amali/Ketua Pentaksir/Assessor dikehendaki mengisi borang ini dan menyerahkannya kepada Penyelia Kawasan untuk dihantar melalui e-mel pada hari yang sama kepada Pengarah Peperiksaan, Lembaga Peperiksaan (u.p.: Unit Menengah Atas/Unit Pra Universiti/Unit Vokasional) dan Pengarah Pendidikan Negeri, Jabatan Pendidikan Negeri (u.p.: Sektor Pentaksiran dan Peperiksaan)."""
        pdf.multi_cell(277, 3.0, arahan, align="J")

        y1 = 48
        pdf.set_xy(10, y1)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(38, 10, "NO. PUSAT", border=1, align="C")
        pdf.cell(90, 10, "NAMA PUSAT PEPERIKSAAN", border=1, align="C")
        x3 = 138
        pdf.set_xy(x3, y1)
        pdf.cell(149, 10, "", border=1)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_xy(x3, y1)
        pdf.cell(149, 3.5, "NAMA KETUA PENGAWAS", align="C")
        pdf.set_xy(x3, y1+3.5)
        pdf.cell(149, 3.5, "PEPERIKSAAN/KETUA PENGAWAS", align="C")
        pdf.set_xy(x3, y1+7)
        pdf.cell(149, 3, "AMALI/KETUA PENTAKSIR/ASSESSOR", align="C")

        pdf.set_xy(10, y1+10)
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(38, 10, no_pusat, border=1, align="C")
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(90, 10, nama_pusat[:48], border=1, align="C")
        pdf.cell(149, 10, nama_kp[:78], border=1, align="C")

        y2 = y1 + 24
        pdf.set_xy(10, y2)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(15, 10, "BIL.", border=1, align="C")
        pdf.cell(38, 10, "ANGKA GILIRAN", border=1, align="C")
        pdf.set_xy(63, y2)
        pdf.cell(42, 10, "", border=1)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_xy(63, y2)
        pdf.cell(42, 5, "NO. PENGENALAN", align="C")
        pdf.set_xy(63, y2+5)
        pdf.cell(42, 5, "DIRI", align="C")
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_xy(105, y2)
        pdf.cell(91, 10, "KES", border=1, align="C")
        pdf.set_xy(196, y2)
        pdf.cell(91, 10, "TINDAKAN", border=1, align="C")

        big_h = 38
        pdf.set_xy(10, y2+10)
        if tiada or not kes_data:
            pdf.cell(15, big_h, "", border=1)
            pdf.cell(38, big_h, "", border=1)
            pdf.cell(42, big_h, "", border=1)
            pdf.cell(91, big_h, "", border=1)
            pdf.cell(91, big_h, "", border=1)
        else:
            pdf.cell(15, big_h, "", border=1)
            pdf.cell(38, big_h, "", border=1)
            pdf.cell(42, big_h, "", border=1)
            pdf.cell(91, big_h, "", border=1)
            pdf.cell(91, big_h, "", border=1)
            pdf.set_font("Helvetica", "", 7.5)
            for idx, k in enumerate(kes_data):
                yy = y2+10 + 1 + idx*6
                if yy > y2+10+big_h-3:
                    break
                pdf.set_xy(10, yy)
                pdf.cell(15, 5, str(idx+1), align="C")
                pdf.set_xy(25, yy)
                pdf.cell(38, 5, str(k.get('angka_giliran',''))[:18], align="C")
                pdf.set_xy(63, yy)
                pdf.cell(42, 5, str(k.get('no_ic',''))[:20], align="C")
                pdf.set_xy(105, yy)
                pdf.cell(91, 5, str(k.get('kes',''))[:55], align="L")
                pdf.set_xy(196, yy)
                pdf.cell(91, 5, str(k.get('tindakan',''))[:55], align="L")

        yf = y2+10+big_h+4
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_xy(10, yf)
        pdf.cell(60, 5, "Pelapor:")
        pdf.set_xy(160, yf)
        pdf.cell(60, 5, "Pengesahan:")
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_xy(10, yf+10)
        pdf.cell(80, 4, "Tandatangan Ketua Pengawas Peperiksaan/")
        pdf.set_xy(10, yf+14)
        pdf.cell(80, 4, "Ketua Pengawas Amali/Ketua Pentaksir/Assessor")
        pdf.set_xy(160, yf+10)
        pdf.cell(80, 4, "Tandatangan Penyelia Kawasan/Pegawai Peperiksaan")
        pdf.line(10, yf+18, 90, yf+18)
        pdf.line(160, yf+18, 240, yf+18)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_xy(10, yf+20)
        pdf.cell(80, 5, "Tarikh: ______________________")
        pdf.set_xy(160, yf+20)
        pdf.cell(80, 5, "Tarikh: ______________________")

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



