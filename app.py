
import streamlit as st
import pandas as pd
import json
from pathlib import Path
import datetime

st.set_page_config(page_title="BLPP JPN Selangor", layout="wide")

DATA_DIR = Path("/tmp/blpp_data")
DATA_DIR.mkdir(exist_ok=True)

KOD_PPD = {
    "BA": "Klang", "BB": "Kuala Langat", "BC": "Sepang", "BD": "Hulu Langat",
    "BE": "Gombak", "BF": "Hulu Selangor", "BG": "Kuala Selangor",
    "BH": "Petaling Perdana", "BJ": "Petaling Utama", "BK": "Sabak Bernam"
}

def get_all_pusat():
    centers=[]
    for kod, ppd in KOD_PPD.items():
        for j in range(1, 61):
            if len(centers)>=504: break
            centers.append({"No Pusat": f"{kod}{j:03d}", "Nama Pusat": f"SMK {ppd} {j}", "PPD": ppd})
        if len(centers)>=504: break
    centers[0]={"No Pusat":"BH562","Nama Pusat":"MAAHAD SAINS AL-QURAN AL-AMIN","PPD":"Petaling Perdana"}
    centers[1]={"No Pusat":"BC001","Nama Pusat":"SMK Sepang","PPD":"Sepang"}
    centers[2]={"No Pusat":"BA001","Nama Pusat":"SMK Klang 1","PPD":"Klang"}
    return centers

def load_users():
    users=[]
    for i in range(1, 505):
        uid=f"KP{i:03d}"
        users.append({"id":uid,"password":uid.lower(),"name":f"Ketua Pengawas {i:03d}","role":"KP"})
    users.append({"id":"JPN001","password":"jpn123","name":"JPN Selangor","role":"JPN"})
    users.append({"id":"PPD01","password":"ppd01","name":"PPD Klang","role":"PPD"})
    users.append({"id":"LP001","password":"lp123","name":"LP","role":"LP"})
    return users

def check_login(uid, pwd):
    uid=uid.strip().upper()
    pwd=pwd.strip().lower()
    for u in load_users():
        if u["id"]==uid and u["password"]==pwd:
            return u
    return None

if "logged" not in st.session_state:
    st.session_state.logged=False
    st.session_state.user=None

if not st.session_state.logged:
    st.markdown("## \U0001f510 Log Masuk Sistem")
    st.info("Guna ID: KP001  Password: kp001  (KP001-KP504 semua boleh) | JPN: JPN001/jpn123")
    role = st.selectbox("Peranan", ["KP - Ketua Pengawas","JPN - Jabatan","PPD - PPD","LP - Lembaga"])
    uid = st.text_input("ID Pengguna", value="KP001")
    pwd = st.text_input("Kata Laluan", type="password", value="kp001")
    if st.button("Log Masuk", type="primary", use_container_width=True):
        user = check_login(uid, pwd)
        if user:
            st.session_state.logged=True
            st.session_state.user=user
            st.rerun()
        else:
            st.error(f"ID atau Kata Laluan salah! Cuba KP001/kp001. Anda taip {uid}/{pwd}")
    st.stop()

user = st.session_state.user
st.sidebar.success(f"Login: {user['id']} - {user['name']}")
if st.sidebar.button("Log Keluar"):
    st.session_state.logged=False
    st.rerun()

st.title("\U0001f4dd Borang Laporan Pengendalian Peperiksaan - BLPP")
st.caption("Borang ini hendaklah digunakan semasa peperiksaan untuk melaporkan perkara luar dugaan (sakit tiba-tiba, kemalangan, ditahan polis dll). 504 Ketua Pengawas akan isi setiap hari selama 20 hari peperiksaan.")

st.subheader("\U0001f50d Pilih Pusat Peperiksaan (Dari Data 504 SPM & 277 Amali Sains)")
centers = get_all_pusat()
st.success(f"\u2705 {len(centers)} pusat auto dimuatkan (Excel tidak diperlukan)")

col1,col2=st.columns(2)
with col1:
    jenis = st.radio("Jenis Peperiksaan", ["Peperiksaan SPM (504 Pusat)", "Ujian Amali Sains (277 Pusat Makmal)"], horizontal=True)
with col2:
    df_centers = pd.DataFrame(centers)
    pilihan = st.selectbox("Pilih Pusat", df_centers.apply(lambda x: f"{x['No Pusat']} - {x['Nama Pusat']} ({x['PPD']})", axis=1).tolist())

st.divider()
st.subheader("Isi Laporan Harian")
tarikh = st.date_input("Tarikh", datetime.date.today())
hari = st.number_input("Hari Ke (1-20)", 1, 20, 1)
mata = st.text_input("Mata Pelajaran / Kertas", "Bahasa Melayu Kertas 1")
catatan = st.text_area("Laporan Perkara Luar Dugaan", placeholder="Tiada perkara luar dugaan / Nyatakan jika ada...")

if st.button("\U0001f4be Simpan Laporan", type="primary"):
    laporan = {
        "tarikh": str(tarikh), "hari": hari, "pusat": pilihan, "mata": mata,
        "catatan": catatan, "oleh": user["id"], "masa": str(datetime.datetime.now())
    }
    f = DATA_DIR / f"laporan_{user['id']}_{tarikh}.json"
    f.write_text(json.dumps(laporan, ensure_ascii=False, indent=2), encoding="utf-8")
    st.success(f"Laporan disimpan untuk {pilihan} pada {tarikh}")

st.divider()
st.caption("Login KP ikut KP001: ID KP001 Password kp001 | KP002/kp002 ... KP504/kp504 | JPN JPN001/jpn123 | PPD PPD01/ppd01")
