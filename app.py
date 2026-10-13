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
    """PDF cantik kemas - 3 layer fallback, takkan return kosong"""
    no_pusat = report.get('no_pusat','')
    nama_pusat = report.get('nama_pusat','')
    nama_kp = report.get('nama_kp','')
    pep = report.get('peperiksaan','')
    thn = report.get('tahun','')
    kes_data = report.get('kes_data', [])
    tiada = report.get('tiada_kes', False)

    # LAYER 1: FPDF cantik dengan logo
    try:
        from fpdf import FPDF
        from PIL import Image
        import os

        logo_to_use = None
        for lp in ["logo_lp.png", "data/logo_lp.png", "./logo_lp.png", "/tmp/logo_lp_final.png", "/mnt/data/logo_lp.png"]:
            if os.path.exists(lp):
                try:
                    im = Image.open(lp)
                    tmp = "/tmp/logo_lp_final.png"
                    im.convert("RGBA").save(tmp, "PNG")
                    logo_to_use = tmp
                    break
                except:
                    try:
                        # cuba guna terus tanpa convert
                        logo_to_use = lp
                        break
                    except:
                        pass

        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=False)
        pdf.add_page()

        if logo_to_use:
            try:
                pdf.image(logo_to_use, x=10, y=4, w=32)
            except Exception as e:
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
        pep_text = f"PEPERIKSAAN {pep} TAHUN {thn}" if pep and thn else "PEPERIKSAAN ........................................ TAHUN ........................................"
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_xy(10, 21)
        pdf.cell(277, 6, pep_text, align="C")
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_xy(10, 27)
        pdf.cell(277, 6, "BORANG LAPORAN PENGENDALIAN PEPERIKSAAN", align="C")

        pdf.set_font("Helvetica", "", 6.8)
        pdf.set_xy(10, 35)
        arahan = """1. Borang ini hendaklah digunakan semasa peperiksaan untuk melaporkan mengenai perkara-perkara di luar dugaan yang berlaku kepada calon seperti sakit secara tiba-tiba, kemalangan, ditahan polis dan lain-lain.
2. Jika ada kes, Ketua Pengawas Peperiksaan/Ketua Pengawas Amali/Ketua Pentaksir/Assessor dikehendaki mengisi borang ini dan menyerahkannya kepada Penyelia Kawasan untuk dihantar melalui e-mel pada hari yang sama kepada Pengarah Peperiksaan, Lembaga Peperiksaan (u.p.: Unit Menengah Atas/Unit Pra Universiti/Unit Vokasional) dan Pengarah Pendidikan Negeri, Jabatan Pendidikan Negeri (u.p.: Sektor Pentaksiran dan Peperiksaan)."""
        pdf.multi_cell(277, 3.2, arahan, align="J")

        y1 = 47
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_xy(10, y1)
        pdf.cell(38, 8, "NO. PUSAT", border=1, align="C")
        pdf.cell(90, 8, "NAMA PUSAT PEPERIKSAAN", border=1, align="C")
        pdf.cell(149, 8, "NAMA KETUA PENGAWAS PEPERIKSAAN/KETUA PENGAWAS AMALI/KETUA PENTAKSIR/ASSESSOR", border=1, align="C")
        pdf.set_xy(10, y1+8)
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(38, 8, no_pusat, border=1, align="C")
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(90, 8, nama_pusat[:48], border=1, align="C")
        pdf.cell(149, 8, nama_kp[:78], border=1, align="C")

        y2 = y1 + 20
        pdf.set_xy(10, y2)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(15, 7, "BIL.", border=1, align="C")
        pdf.cell(38, 7, "ANGKA GILIRAN", border=1, align="C")
        pdf.cell(42, 7, "NO. PENGENALAN DIRI", border=1, align="C")
        pdf.cell(91, 7, "KES", border=1, align="C")
        pdf.cell(91, 7, "TINDAKAN", border=1, align="C")

        pdf.set_font("Helvetica", "", 8.5)
        if tiada or not kes_data:
            for i in range(6):
                yy = y2 + 7 + (i*7)
                pdf.set_xy(10, yy)
                pdf.cell(15, 7, "", border=1, align="C")
                pdf.cell(38, 7, "", border=1, align="C")
                pdf.cell(42, 7, "", border=1, align="C")
                pdf.cell(91, 7, "", border=1, align="L")
                pdf.cell(91, 7, "", border=1, align="L")
        else:
            for i in range(6):
                yy = y2 + 7 + (i*7)
                pdf.set_xy(10, yy)
                if i < len(kes_data):
                    k = kes_data[i]
                    pdf.cell(15, 7, str(i+1), border=1, align="C")
                    pdf.cell(38, 7, str(k.get('angka_giliran',''))[:18], border=1, align="C")
                    pdf.cell(42, 7, str(k.get('no_ic',''))[:20], border=1, align="C")
                    pdf.cell(91, 7, str(k.get('kes',''))[:60], border=1, align="L")
                    pdf.cell(91, 7, str(k.get('tindakan',''))[:60], border=1, align="L")
                else:
                    pdf.cell(15, 7, "", border=1, align="C")
                    pdf.cell(38, 7, "", border=1, align="C")
                    pdf.cell(42, 7, "", border=1, align="C")
                    pdf.cell(91, 7, "", border=1, align="L")
                    pdf.cell(91, 7, "", border=1, align="L")

        yf = y2 + 7 + 42 + 6
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_xy(10, yf)
        pdf.cell(60, 5, "Pelapor:")
        pdf.set_xy(160, yf)
        pdf.cell(60, 5, "Pengesahan:")
        pdf.set_font("Helvetica", "", 8)
        pdf.set_xy(10, yf+12)
        pdf.cell(70, 4, "Tandatangan Ketua Pengawas Peperiksaan/")
        pdf.set_xy(10, yf+16)
        pdf.cell(70, 4, "Ketua Pengawas Amali/Ketua Pentaksir/Assessor")
        pdf.set_xy(160, yf+12)
        pdf.cell(70, 4, "Tandatangan Penyelia Kawasan/Pegawai Peperiksaan")
        pdf.line(10, yf+20, 85, yf+20)
        pdf.line(160, yf+20, 235, yf+20)
        pdf.set_xy(10, yf+22)
        pdf.cell(80, 5, "Tarikh: ___________________________")
        pdf.set_xy(160, yf+22)
        pdf.cell(80, 5, "Tarikh: ___________________________")

        out = pdf.output(dest='S')
        if isinstance(out, str):
            return out.encode('latin1')
        return bytes(out)
    except Exception as e:
        import traceback
        traceback.print_exc()

    # LAYER 2: Reportlab simple tanpa logo - pasti jadi
    try:
        from io import BytesIO
        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT

        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
        story = []
        def P(txt, sz=9, bold=False):
            return Paragraph(f"<b>{txt}</b>" if bold else txt, ParagraphStyle(f"s{sz}{bold}{txt[:5]}", fontName='Helvetica-Bold' if bold else 'Helvetica', fontSize=sz, leading=sz+2, alignment=TA_CENTER))

        W = landscape(A4)[0] - 20*mm
        hdr = [[P("LEMBAGA PEPERIKSAAN<br/>KEMENTERIAN PENDIDIKAN MALAYSIA",10,True), P("LP/Am 90 Pin. 1/2023<br/>BLPP",9,True)]]
        story.append(Table(hdr, colWidths=[0.8*W,0.2*W]))
        story.append(Spacer(1,3*mm))
        story.append(P(f"PEPERIKSAAN {pep} TAHUN {thn}" if pep and thn else "PEPERIKSAAN ... TAHUN ...",11,True))
        story.append(P("BORANG LAPORAN PENGENDALIAN PEPERIKSAAN",12,True))
        story.append(Spacer(1,3*mm))

        t1 = [[P("NO. PUSAT",8,True), P("NAMA PUSAT",8,True), P("NAMA KETUA PENGAWAS",8,True)],
              [P(no_pusat,9,True), P(nama_pusat,9,True), P(nama_kp,9,True)]]
        tbl1 = Table(t1, colWidths=[W*0.14,W*0.34,W*0.52])
        tbl1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.6,colors.black)]))
        story.append(tbl1)
        story.append(Spacer(1,3*mm))

        hdr2 = [P("BIL.",8,True), P("ANGKA GILIRAN",8,True), P("NO. PENGENALAN DIRI",8,True), P("KES",8,True), P("TINDAKAN",8,True)]
        # SEBIJIK ASAL: 1 header + 1 kotak besar
        big_row = [P("",8) for _ in range(5)]
        if tiada or not kes_data:
            data2 = [hdr2, big_row]
        else:
            # Isi dalam satu baris besar
            bil_txt = "<br/>".join([str(i+1) for i in range(len(kes_data))])
            ang_txt = "<br/>".join([k.get('angka_giliran','') for k in kes_data])
            ic_txt = "<br/>".join([k.get('no_ic','') for k in kes_data])
            kes_txt = "<br/>".join([k.get('kes','') for k in kes_data])
            tind_txt = "<br/>".join([k.get('tindakan','') for k in kes_data])
            big_row = [P(bil_txt,8), P(ang_txt,8), P(ic_txt,8), P(kes_txt,8), P(tind_txt,8)]
            data2 = [hdr2, big_row]

        tbl2 = Table(data2, colWidths=[W*0.06,W*0.15,W*0.17,W*0.31,W*0.31])
        tbl2.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.6,colors.black)]))
        story.append(tbl2)

        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        import traceback
        traceback.print_exc()

    # LAYER 3: Paling last - minimal PDF
    try:
        from fpdf import FPDF
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        pdf.set_font("Helvetica","B",14)
        pdf.cell(0,10,f"BLPP - {no_pusat} - {nama_pusat}", align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica","",11)
        pdf.cell(0,8,f"Peperiksaan: {pep} Tahun: {thn}", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0,8,f"Nama KP: {nama_kp}", new_x="LMARGIN", new_y="NEXT")
        out = pdf.output(dest='S')
        return out.encode('latin1') if isinstance(out,str) else bytes(out)
    except:
        return f"BLPP {no_pusat} {nama_pusat}".encode()





# ========== LOGIN ==========
# ========== LOGIN ==========
# ========== LOGIN ==========

if "login" not in st.session_state:
    st.session_state.login = False
    st.session_state.user = None

if not st.session_state.login:
    # FORCE CLEAN OLD CACHE - penting untuk KP001
    import os
    try:
        if os.path.exists(USERS_FILE):
            os.remove(USERS_FILE)
    except:
        pass
    
    st.markdown('<div class="header-biru"><h1>📋 SISTEM PENGURUSAN PUSAT PEPERIKSAAN</h1><h2>BLPP - LP/Am 90 Pin. 1/2023</h2><h3>JPN SELANGOR - 504 PUSAT</h3></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1,1.2,1])
    with col2:
        st.markdown('<div class="card"><h3 style="text-align:center; margin:0 0 16px 0;">🔐 Log Masuk Sistem</h3>', unsafe_allow_html=True)
        users = load_users()
        peranan = st.selectbox("Peranan", ["KP - Ketua Pengawas", "PPD", "ADMIN_JPN - JPN Selangor", "LP - Lembaga Peperiksaan"])
        id_pengguna = st.text_input("ID Pengguna", placeholder="Contoh: KP001 / PPD01 / JPN001 / LP001")
        password = st.text_input("Kata Laluan", type="password")
        
        if st.button("Log Masuk", type="primary", use_container_width=True):
            found = None
            uid_input = id_pengguna.strip().upper()
            pwd_input = password.strip().lower()  # FIX: lowercase password
            for u in users:
                # FIX: password check lowercase
                if u["id"].upper() == uid_input and str(u["password"]).lower() == pwd_input:
                    if peranan.startswith(u["peranan"]) or u["peranan"] in peranan:
                        found = u
                        break
            if found:
                st.session_state.login = True
                st.session_state.user = found
                st.success(f"Selamat datang {found['nama']}!")
                st.rerun()
            else:
                st.error(f"ID atau Kata Laluan salah! Cuba: KP001 / kp001 (Debug: {len(users)} users loaded)")
        
        st.markdown('<div style="margin-top:14px; padding:10px; background:#F0F7FF; border-radius:8px; font-size:11px;"><b>Demo Login:</b><br/>KP: KP001 / kp001 (504 pusat)<br/>PPD: PPD01 / ppd01 (10 PPD)<br/>JPN: JPN001 / jpn123<br/>LP: LP001 / lp123<br/>BH562 / bh562 juga boleh</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ========== DASHBOARD UTAMA ==========
user = st.session_state.user
peranan_user = user["peranan"]

st.markdown(f'<div class="header-biru"><h1>📋 SISTEM BLPP - {user["nama"]}</h1><h2>Peranan: {peranan_user} | PPD: {user.get("ppd","-")} | No Pusat: {user.get("no_pusat","-")}</h2><h3>BLPP LP/Am 90 Pin. 1/2023 - Laporan Harian Pusat Peperiksaan</h3></div>', unsafe_allow_html=True)

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
    st.markdown(f'''
    <div class="card" style="position:sticky; top:18px; padding:16px;">
        <div style="font-weight:800; font-size:14px;">{user['nama']}</div>
        <div style="font-size:11px; color:#64748B; margin:4px 0;">{peranan_user}</div>
        <div style="font-size:11px; background:#EEF2FF; padding:4px 8px; border-radius:6px; display:inline-block;">PPD: {user.get('ppd','-')}</div>
        <hr style="margin:12px 0;">
    ''', unsafe_allow_html=True)
    
    menu = st.radio("Menu", 
        ["📝 Isi Borang BLPP (KP)" if peranan_user=="KP" else "📊 Dashboard",
         "📋 Senarai Laporan",
         "🖨️ Cetak Borang BLPP",
         "👥 Pengurusan Pengguna (JPN)" if peranan_user=="ADMIN_JPN" else "📈 Statistik PPD" if peranan_user=="PPD" else "📈 Statistik LP" if peranan_user=="LP" else "📈 Statistik Saya",
         "🚪 Log Keluar"
        ], label_visibility="collapsed")
    
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
            st.markdown("#### 🔍 Pilih Pusat Peperiksaan (Dari Data 504 SPM & 277 Amali Sains)")
            # Step 1: Jenis peperiksaan
            jenis_pep = st.radio("Jenis Peperiksaan", JENIS_PEPERIKSAAN, horizontal=True)
        
            # Load centers based on jenis - FIX AUTO GENERATE tanpa Excel
            if "Amali Sains" in jenis_pep:
                centers = load_amali_centers()
                exam_mode = "AMALI_SAINS"
            else:
                centers = load_spm_centers()
                exam_mode = "SPM"
        
            # AUTO FALLBACK jika Excel tak ada
            if not centers:
                # generate 504 auto
                try:
                    centers = get_all_pusat()
                    if centers:
                        st.success(f"✅ {len(centers)} pusat auto dimuatkan (Excel tidak diperlukan) - Mode {exam_mode}")
                except:
                    centers = []
            
            if not centers:
                # last fallback - buat 504 dummy
                centers = []
                for kod, ppd in KOD_PPD_MAPPING.items():
                    for j in range(1, 60):
                        if len(centers)>=504:
                            break
                        centers.append({"Kod_PPD": kod, "No_Pusat": f"{kod}{j:03d}", "Nama_Pusat": f"SMK {ppd} {j}", "PPD": ppd, "Nama_Sekolah": f"SMK {ppd} {j}"})
                    if len(centers)>=504:
                        break
                # pastikan BH562 ada
                if centers:
                    centers[0] = {"Kod_PPD":"BH","No_Pusat":"BH562","Nama_Pusat":"MAAHAD SAINS AL-QURAN AL-AMIN","PPD":"Petaling Perdana","Nama_Sekolah":"MAAHAD SAINS AL-QURAN AL-AMIN"}
                st.warning(f"⚠️ Guna data auto {len(centers)} pusat (Excel asal tidak ditemui)")

            # ensure tidak kosong
            centers = ensure_centers(centers) if 'ensure_centers' in globals() else centers
        
            # Step 2: Pilih Daerah
            colD1, colD2, colD3 = st.columns([1,1,1])
            with colD1:
                # Filter ikut PPD user jika KP, tapi boleh pilih lain juga
                selected_ppd = st.selectbox("Pilih Daerah / PPD", PPD_LIST, index=PPD_LIST.index(user["ppd"]) if user["ppd"] in PPD_LIST else 0)
        
            # Filter centers by PPD
            filtered_by_ppd = [c for c in centers if c["PPD"] == selected_ppd]
        
            with colD2:
                if filtered_by_ppd:
                    # Pilih Nama Sekolah
                    sekolah_options = sorted(list(set([c["Nama_Pusat"] for c in filtered_by_ppd])))
                    selected_sekolah = st.selectbox(f"Pilih Nama Sekolah ({len(sekolah_options)} sekolah di {selected_ppd})", sekolah_options)
                    # Cari No Pusat dari sekolah
                    matched = [c for c in filtered_by_ppd if c["Nama_Pusat"] == selected_sekolah]
                    if matched:
                        auto_no_pusat = matched[0]["No_Pusat"]
                        auto_kod_ppd = matched[0]["Kod_PPD"]
                    else:
                        auto_no_pusat = ""
                        auto_kod_ppd = PPD_TO_KOD.get(selected_ppd, "")
                else:
                    st.warning(f"Tiada pusat di {selected_ppd} untuk {jenis_pep}")
                    selected_sekolah = ""
                    auto_no_pusat = ""
                    auto_kod_ppd = ""
        
            with colD3:
                st.markdown("**Maklumat Pusat (Auto)**")
                if auto_no_pusat:
                    st.success(f"No Pusat: {auto_no_pusat} | Kod: {auto_kod_ppd} | {selected_ppd}")
                else:
                    st.info("Pilih sekolah untuk auto No Pusat")
        
            # Jika dah pilih sekolah, baru keluar form BLPP
            if not auto_no_pusat or not selected_sekolah:
                st.info("👆 Sila pilih Daerah dan Nama Sekolah untuk papar borang BLPP")
                st.stop()
        
            # Form BLPP sebenar
            st.markdown("---")
            st.markdown(f"### 📝 Borang BLPP - {selected_sekolah} ({auto_no_pusat}) - {selected_ppd}")
        
            with st.form("blpp_form", clear_on_submit=False):
                colA, colB, colC = st.columns(3)
                with colA:
                    peperiksaan = st.selectbox("Peperiksaan", PEPERIKSAAN_LIST, index=0 if exam_mode=="SPM" else 4)
                    tahun = st.text_input("Tahun", value="2026")
                    no_pusat = st.text_input("No. Pusat (Auto dari Data)", value=auto_no_pusat, disabled=False)
                with colB:
                    nama_pusat = st.text_input("Nama Pusat Peperiksaan (Auto)", value=selected_sekolah)
                    ppd = st.text_input("PPD / Daerah (Auto)", value=selected_ppd)
                    kod_ppd_display = st.text_input("Kod PPD", value=auto_kod_ppd, disabled=True)
                with colC:
                    tarikh = st.date_input("Tarikh Laporan", value=date.today())
                    nama_kp = st.text_input("Nama KP", value=user["nama"])
                    # Simpan exam_mode untuk backend
                    st.text_input("Jenis Peperiksaan (System)", value=exam_mode, disabled=True, key="exam_mode_hidden")

            
                st.markdown("---")
                st.markdown("#### 📋 Butiran Kes Luar Dugaan")
                tiada_kes = st.checkbox("☑️ TIADA KES LUAR DUGAAN PADA HARI INI (Tick jika tiada kes)", value=False)
            
                kes_data = []
                if not tiada_kes:
                    st.markdown("**Isi maklumat calon yang terlibat (boleh tambah banyak):**")
                    num_kes = st.number_input("Bilangan Kes", min_value=1, max_value=20, value=1)
                    for i in range(num_kes):
                        st.markdown(f"**Kes {i+1}**")
                        c1, c2, c3 = st.columns([1,1,2])
                        with c1:
                            ag = st.text_input(f"Angka Giliran {i+1}", key=f"ag_{i}", placeholder="Contoh: BC123456")
                        with c2:
                            ic = st.text_input(f"No. IC {i+1}", key=f"ic_{i}", placeholder="990101-10-1234")
                        with c3:
                            pass
                        c4, c5 = st.columns(2)
                        with c4:
                            kes = st.text_area(f"Kes {i+1} (sakit tiba-tiba / kemalangan / ditahan polis / lain-lain)", key=f"kes_{i}", height=80, placeholder="Nyatakan kes secara terperinci...")
                        with c5:
                            tindakan = st.text_area(f"Tindakan {i+1}", key=f"tind_{i}", height=80, placeholder="Tindakan yang diambil...")
                        kes_data.append({"angka_giliran": ag, "no_ic": ic, "kes": kes, "tindakan": tindakan})
            
                st.markdown("---")
                st.markdown("### ✍️ Tandatangan Digital & Pengesahan Pelapor")
                colT1, colT2 = st.columns([1,1])
                with colT1:
                    pelapor_nama = st.text_input("Nama Pelapor (KP)", value=user["nama"])
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
                    tt_b64 = ""
                    if tandatangan_file:
                        img = Image.open(tandatangan_file)
                        # Compress to <100KB
                        img.thumbnail((500, 300))
                        buf = BytesIO()
                        img.save(buf, format="PNG", optimize=True)
                        tt_b64 = base64.b64encode(buf.getvalue()).decode()
                        st.image(img, caption="Tandatangan Diupload", width=250)
                        st.success(f"✅ Tandatangan: {len(tt_b64)/1024:.1f}KB")
                    elif camera_tt:
                        img = Image.open(camera_tt)
                        img.thumbnail((500, 300))
                        buf = BytesIO()
                        img.save(buf, format="PNG", optimize=True)
                        tt_b64 = base64.b64encode(buf.getvalue()).decode()
                        st.image(img, caption="Tandatangan dari Kamera", width=250)
                    else:
                        st.info("Belum ada tandatangan. Sila upload atau guna kamera.")
                        # Lukis manual guna canvas simple - guna text input sebagai alternatif
                        st.markdown("Atau taip nama sebagai tandatangan digital:")
                        tt_text = st.text_input("Tandatangan Text", value=user["nama"], key="tt_text")
                        if tt_text:
                            # Generate simple signature image from text
                            from PIL import ImageDraw, ImageFont
                            sig_img = Image.new('RGB', (400, 100), color='white')
                            d = ImageDraw.Draw(sig_img)
                            try:
                                d.text((10,30), tt_text, fill=(13,71,161), font=ImageFont.load_default())
                            except:
                                d.text((10,30), tt_text, fill=(13,71,161))
                            buf = BytesIO()
                            sig_img.save(buf, format="PNG")
                            tt_b64 = base64.b64encode(buf.getvalue()).decode()
                            st.image(sig_img, width=250)

            
                submitted = st.form_submit_button("📤 Hantar Laporan BLPP", type="primary", use_container_width=True)
            
                if submitted:
                    if not no_pusat or not nama_pusat:
                        st.error("Sila isi No. Pusat dan Nama Pusat!")
                    else:
                        new_report = {
                            "id": f"BLPP-{datetime.now().strftime('%Y%m%d%H%M%S')}-{no_pusat}",
                            "peperiksaan": peperiksaan,
                            "tahun": tahun,
                            "no_pusat": no_pusat,
                            "nama_pusat": nama_pusat,
                            "nama_kp": nama_kp,
                            "ppd": ppd,
                            "tarikh": tarikh.strftime("%Y-%m-%d"),
                            "tarikh_submit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "kes_data": [] if tiada_kes else kes_data,
                            "tiada_kes": tiada_kes,
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
                st.markdown("### 📄 Borang BLPP Lengkap - TELAH DIISI DATA (Bukan Kosong!)")
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
            st.markdown("#### 📜 Laporan Saya (20 Hari Peperiksaan)")
            my_reports = [r for r in reports if r.get("id_kp")==user["id"]]
            if my_reports:
                df_my = pd.DataFrame(my_reports)
                st.dataframe(df_my[["id","tarikh","peperiksaan","no_pusat","tiada_kes"]], use_container_width=True)
            else:
                st.info("Belum ada laporan. Sila isi borang di atas untuk hari ini.")

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

        filtered = reports
        if peranan_user == "PPD":
            filtered = [r for r in filtered if r.get("ppd")==user["ppd"]]
        elif peranan_user == "KP":
            filtered = [r for r in filtered if r.get("id_kp")==user["id"]]
    
        if filter_ppd != "Semua":
            filtered = [r for r in filtered if r.get("ppd")==filter_ppd]
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

    # ========== MENU 3: CETAK BORANG BLPP ==========
    elif menu.startswith("🖨️"):
        st.markdown('<div class="card"><h3>🖨️ Cetak Borang BLPP LP/Am 90 Pin. 1/2023 (Official)</h3><p style="font-size:12px;">Pilih laporan untuk cetak borang rasmi BLPP. JPN boleh pilih PPD mana, LP boleh view ikut PPD.</p></div>', unsafe_allow_html=True)
    
        if not reports:
            st.warning("Belum ada laporan untuk dicetak.")
        else:
            # Filter ikut peranan
            available_reports = reports
            if peranan_user == "PPD":
                available_reports = [r for r in reports if r.get("ppd")==user["ppd"]]
            elif peranan_user == "KP":
                available_reports = [r for r in reports if r.get("id_kp")==user["id"]]
        
            if peranan_user in ["ADMIN_JPN","LP"]:
                sel_ppd = st.selectbox("Pilih PPD untuk Cetak (JPN/LP boleh pilih)", ["Semua"] + PPD_LIST)
                if sel_ppd != "Semua":
                    available_reports = [r for r in available_reports if r.get("ppd")==sel_ppd]
        
            if not available_reports:
                st.info("Tiada laporan untuk PPD tersebut.")
            else:
                report_options = {}
                for r in available_reports:
                    kes_label = "TIADA KES" if r.get('tiada_kes') else f"{len(r.get('kes_data',[]))} KES"
                    label = f"{r['id']} | {r['tarikh']} | {r['no_pusat']} | {r['nama_pusat']} | {r['ppd']} | {kes_label}"
                    report_options[label] = r
                selected_label = st.selectbox("Pilih Laporan untuk Cetak", list(report_options.keys()))
                selected_report = report_options[selected_label]
            
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
                    count = len([r for r in reports if r.get("ppd")==ppd])
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
            # Statistik untuk PPD / LP / KP
            st.markdown(f'<div class="card"><h3>📈 Statistik {peranan_user} - {user.get("ppd","")}</h3></div>', unsafe_allow_html=True)
        
            my_reports = reports
            if peranan_user == "PPD":
                my_reports = [r for r in reports if r.get("ppd")==user["ppd"]]
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
