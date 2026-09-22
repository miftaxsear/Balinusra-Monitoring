import json
import os
import numpy as np
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Balinusra Monitoring", layout="wide")

# CSS Kustom untuk background teks label Pilih Bulan & WSID
st.markdown(
    """
    <style>
        /* Background teks label Pilih Bulan (Light Turquoise) */
        div[data-testid="stSelectbox"]:has(label:contains("Pilih Bulan:")) label p {
            background-color: #AFEEEE !important;
            color: #004D40 !important;
            padding: 2px 8px !important;
            border-radius: 4px !important;
            font-weight: bold !important;
            display: inline-block !important;
        }

        /* Background teks label Pilih / Ketik WSID (Pink) */
        div[data-testid="stSelectbox"]:has(label:contains("Pilih / Ketik WSID:")) label p {
            background-color: #FFB6C1 !important;
            color: #880E4F !important;
            padding: 2px 8px !important;
            border-radius: 4px !important;
            font-weight: bold !important;
            display: inline-block !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# -------------------------------------------------------------
# 2. DATABASE USER LOGIN & SESSION STATE
# -------------------------------------------------------------
USERS = {
    "admin": {
        "password": "admin123",
        "role": "admin",
        "name": "Administrator",
    },
    "user": {"password": "user123", "role": "viewer", "name": "User Biasa"},
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["role"] = None
    st.session_state["user_name"] = ""

# Inisialisasi session state untuk navigasi jika belum ada
if "nav_menu" not in st.session_state:
    st.session_state["nav_menu"] = "Uptime FRX DT (WSID)"


# -------------------------------------------------------------
# 3. HALAMAN LOGIN STYLE MYDATINDO
# -------------------------------------------------------------
def show_login_page():
    login_css = """
    <style>
        header {visibility: hidden;}
        .block-container {
            padding-top: 3rem !important;
            padding-bottom: 2rem !important;
        }
        
        .brand-title {
            font-family: 'Lucida Calligraphy', 'Lucida Handwriting', 'Apple Chancery', cursive;
            color: #27ae60;
            font-size: 34px;
            font-weight: bold;
            text-align: center;
            margin-bottom: 12px;
        }

        .login-subtitle {
            text-align: center;
            color: #666666;
            font-size: 13px;
            margin-bottom: 18px;
        }

        div[data-baseweb="input"] {
            background-color: #eef4fb !important;
            border: 1px solid #b8d3f2 !important;
            border-radius: 3px !important;
        }
        
        div[data-baseweb="input"] input {
            color: #2c3e50 !important;
            font-size: 14px !important;
        }

        div.stButton > button {
            width: 100% !important;
            background-color: #3488b5 !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 3px !important;
            padding: 8px 0px !important;
            font-weight: bold !important;
            font-size: 14px !important;
            margin-top: 8px !important;
        }

        div.stButton > button:hover {
            background-color: #286f95 !important;
            color: #ffffff !important;
        }

        .login-footer {
            margin-top: 22px;
            font-size: 11px;
            color: #777777;
            text-align: center;
        }
    </style>
    """
    st.markdown(login_css, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown(
            '<div class="brand-title">Balinusra Monitoring</div>',
            unsafe_allow_html=True,
        )

        with st.container(border=True):
            st.markdown(
                '<div class="login-subtitle">Sign in to start your session</div>',
                unsafe_allow_html=True,
            )

            st.components.v1.html(
                """
                <script>
                const inputs = window.parent.document.querySelectorAll('input');
                if (inputs.length >= 2) {
                    inputs[0].setAttribute('autocomplete', 'username');
                    inputs[0].setAttribute('name', 'username');
                    inputs[1].setAttribute('autocomplete', 'current-password');
                    inputs[1].setAttribute('name', 'password');
                }
                </script>
                """,
                height=0,
            )

            with st.form("login_form"):
                username_input = st.text_input(
                    "Username",
                    placeholder="Username",
                    label_visibility="collapsed",
                )
                password_input = st.text_input(
                    "Password",
                    type="password",
                    placeholder="Password",
                    label_visibility="collapsed",
                )
                submit_button = st.form_submit_button("Sign In")

                if submit_button:
                    if (
                        username_input in USERS
                        and USERS[username_input]["password"] == password_input
                    ):
                        st.session_state["logged_in"] = True
                        st.session_state["role"] = USERS[username_input][
                            "role"
                        ]
                        st.session_state["user_name"] = USERS[username_input][
                            "name"
                        ]
                        st.success("Login berhasil!")
                        st.rerun()
                    else:
                        st.error("Username atau Password salah!")

            st.markdown(
                '<div class="login-footer">Copyright © 2026 <strong>MiftaxSear</strong>. All rights reserved.</div>',
                unsafe_allow_html=True,
            )


if not st.session_state["logged_in"]:
    show_login_page()
    st.stop()


# -------------------------------------------------------------
# 4. FUNGSI MEMBACA DATA MASTER EXCEL
# -------------------------------------------------------------
@st.cache_data
def load_data(file_path):
    xls = pd.ExcelFile(file_path)

    df_frx_raw = (
        pd.read_excel(xls, "FRX DT")
        if "FRX DT" in xls.sheet_names
        else pd.DataFrame()
    )

    df_site = (
        pd.read_excel(xls, "SITE BLNS")
        if "SITE BLNS" in xls.sheet_names
        else pd.DataFrame()
    )

    df_visit = (
        pd.read_excel(xls, "MainVisit")
        if "MainVisit" in xls.sheet_names
        else pd.DataFrame()
    )

    df_atm_summary = pd.DataFrame()
    if "atm total" in xls.sheet_names:
        try:
            df_atm_summary = pd.read_excel(xls, "atm total", skiprows=5)
        except Exception:
            pass

    daily_sheets = {}
    num_sheets = [s for s in xls.sheet_names if s.isdigit()]
    num_sheets.sort(key=lambda x: int(x))

    for sname in num_sheets:
        try:
            df_day = pd.read_excel(xls, sname, header=5)
            if "WSID" in df_day.columns:
                daily_sheets[int(sname)] = df_day
        except Exception:
            pass

    return (
        df_frx_raw,
        df_site,
        df_visit,
        df_atm_summary,
        daily_sheets,
    )


# -------------------------------------------------------------
# 5. SIDEBAR: NAVIGASI MENU & PANEL ADMIN
# -------------------------------------------------------------
st.sidebar.markdown(
    "<h3 style='font-size: 18px; font-weight: bold; margin-bottom: 10px; white-space: nowrap;'>🔍 Dasbord Monitoring</h3>",
    unsafe_allow_html=True,
)

menu_options_list = [
    "Uptime FRX DT (WSID)",
    "Uptime Harian",
    "Uptime CSE by Tipe Mesin",
    "Uptime PKT by Tipe Mesin",
    "Riwayat Kunjungan (Visit)",
]

# Menggunakan key="nav_menu" langsung untuk sinkronisasi otomatis tanpa bug 2x klik
menu_option = st.sidebar.radio(
    "Pilih Tampilan Dashboard:",
    options=menu_options_list,
    key="nav_menu",
)

st.sidebar.divider()

month_options = [f"2026-{m:02d}" for m in range(1, 13)]

if st.session_state["role"] == "admin":
    st.sidebar.subheader("⚙️ Panel Admin (All Access)")
    admin_upload_month = st.sidebar.selectbox(
        "Upload Master Excel untuk Bulan:",
        options=month_options,
        index=8,
    )

    uploaded_file = st.sidebar.file_uploader(
        f"Upload File Master ({admin_upload_month}):", type=["xlsx"]
    )

    if uploaded_file is not None:
        save_file_name = f"MASTER_{admin_upload_month}.xlsx"
        with open(save_file_name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.cache_data.clear()
        st.sidebar.success(f"File {save_file_name} berhasil diperbarui!")
        st.rerun()
else:
    st.sidebar.info("👁️ **Mode Viewer**: Anda hanya memiliki akses melihat data.")

st.sidebar.divider()

st.sidebar.markdown(f"**Logged in as:** {st.session_state['user_name']}")
st.sidebar.markdown(f"**Role:** `{st.session_state['role'].upper()}`")

if st.sidebar.button("Logout"):
    st.session_state["logged_in"] = False
    st.session_state["role"] = None
    st.rerun()


# =============================================================
# HEADER ATAS: PILIHAN BULAN, WSID, DAN TOMBOL NAVIGASI
# =============================================================
col_btn, col_month, col_wsid = st.columns([1.5, 1, 1.2])

with col_month:
    selected_month = st.selectbox(
        "Pilih Bulan:",
        options=month_options,
        index=8,  # Default 2026-09
        key="main_month_select",
    )

target_excel_file = f"MASTER_{selected_month}.xlsx"

if not os.path.exists(target_excel_file):
    if os.path.exists("MASTER.xlsx"):
        target_excel_file = "MASTER.xlsx"
    else:
        st.error(
            f"File '{target_excel_file}' atau 'MASTER.xlsx' tidak ditemukan di server!"
        )
        st.stop()

(
    df_frx_raw,
    df_site,
    df_visit,
    df_atm_summary,
    daily_sheets,
) = load_data(target_excel_file)

# Mengambil daftar WSID dari data master
list_wsid = []
if not df_site.empty and "ID" in df_site.columns:
    list_wsid = df_site["ID"].dropna().unique().tolist()
elif not df_atm_summary.empty and "WSID" in df_atm_summary.columns:
    list_wsid = df_atm_summary["WSID"].dropna().unique().tolist()

if "ZTR4" not in list_wsid and list_wsid:
    list_wsid.insert(0, "ZTR4")

selected_wsid = "ZTR4"
with col_wsid:
    if list_wsid:
        selected_wsid = st.selectbox(
            "Pilih / Ketik WSID:",
            options=list_wsid,
            index=0,
            key="main_wsid_select",
        )

with col_btn:
    st.write("")  # Menyejajarkan posisi tombol vertikal dengan selectbox
    st.write("")
    if st.button(
        f"🔗 Buka Riwayat Kunjungan ({selected_wsid})",
        type="primary",
        use_container_width=True,
    ):
        st.session_state["nav_menu"] = "Riwayat Kunjungan (Visit)"
        st.rerun()


# =============================================================
# HALAMAN 1: UPTIME FRX DT (PER WSID)
# =============================================================
if menu_option == "Uptime FRX DT (WSID)":

    info_data = {
        "Wsid": selected_wsid,
        "Ws Name": "-",
        "Trx": "-",
        "Serial No": "-",
        "Sn": "-",
        "Address": "-",
        "City": "-",
        "Province": "-",
        "Island": "-",
        "Location": "-",
        "Vip": "0",
        "Mtype": "CRM",
        "Model": "-",
        "Sw Installed": "-",
        "Vendor": "-",
        "Engineer": "-",
    }

    if not df_site.empty and "ID" in df_site.columns:
        match_site = df_site[
            df_site["ID"].astype(str).str.upper() == str(selected_wsid).upper()
        ]
        if not match_site.empty:
            row = match_site.iloc[0]
            info_data["Ws Name"] = str(row.get("WS_NAME", "-"))
            info_data["Serial No"] = str(row.get("SERIAL_NO", "-"))
            info_data["Address"] = str(row.get("ADDRESS", "-"))
            info_data["City"] = str(row.get("CITY", "-"))
            info_data["Province"] = str(row.get("PROVINCE", "-"))
            info_data["Island"] = str(row.get("ISLAND", "-"))
            info_data["Location"] = str(row.get("LOCATION", "-"))
            info_data["Mtype"] = str(row.get("MTYPE", "CRM")).upper()
            info_data["Model"] = str(row.get("MODEL", "-"))
            info_data["Vendor"] = str(row.get("VENDOR", "-"))
            info_data["Engineer"] = str(row.get("SE_NAME", "-"))

    if not df_frx_raw.empty and len(df_frx_raw) >= 5:
        row_frx_header = df_frx_raw.iloc[4]
        trx_val = row_frx_header.iloc[9]
        if pd.notnull(trx_val) and str(trx_val).strip() not in [
            "",
            "nan",
            "None",
        ]:
            info_data["Trx"] = str(trx_val).strip()

        sn_val = row_frx_header.iloc[10]
        if pd.notnull(sn_val) and str(sn_val).strip() not in ["", "nan", "None"]:
            info_data["Sn"] = str(sn_val).strip()

        if info_data["Engineer"] in ["-", "nan", "None", ""]:
            eng_val = row_frx_header.iloc[1]
            if pd.notnull(eng_val):
                info_data["Engineer"] = str(eng_val).strip()

    if info_data["Trx"] in ["-", "nan", "None", ""]:
        if not df_atm_summary.empty and "WSID" in df_atm_summary.columns:
            match_atm = df_atm_summary[
                df_atm_summary["WSID"].astype(str).str.upper()
                == str(selected_wsid).upper()
            ]
            if not match_atm.empty:
                trx_atm = match_atm.iloc[0].get("TRX", "-")
                if pd.notnull(trx_atm):
                    info_data["Trx"] = str(trx_atm).strip()

    freq_dt_val = 0
    achieve_ut_float = 0.0

    if not df_atm_summary.empty and "WSID" in df_atm_summary.columns:
        match_wsid = df_atm_summary[
            df_atm_summary["WSID"].astype(str).str.upper()
            == str(selected_wsid).upper()
        ]
        if not match_wsid.empty:
            row_atm = match_wsid.iloc[0]
            if "FRX" in row_atm and pd.notnull(row_atm["FRX"]):
                freq_dt_val = int(float(row_atm["FRX"]))
            if "UPTIME" in row_atm and pd.notnull(row_atm["UPTIME"]):
                try:
                    achieve_ut_float = float(row_atm["UPTIME"])
                    if achieve_ut_float <= 1.0:
                        achieve_ut_float = achieve_ut_float * 100
                except Exception:
                    pass

    achieve_ut_str = f"{round(achieve_ut_float, 2)}%"
    mtype = info_data["Mtype"]
    target_ut = 99.75 if "ATM" in mtype else 99.20
    is_tercapai = achieve_ut_float >= target_ut

    status_text = "Tercapai" if is_tercapai else "Tidak tercapai"
    status_color = "#27ae60" if is_tercapai else "#e74c3c"
    ut_color = status_color
    frx_color = "#e74c3c" if freq_dt_val > 0 else "#27ae60"

    chart_dates = []
    chart_uptime = []
    table_frx_rows_html = []
    year_str, month_str = selected_month.split("-")

    if daily_sheets:
        sorted_days = sorted(daily_sheets.keys())
        for d in sorted_days:
            df_d = daily_sheets[d]
            row_match = df_d[
                df_d["WSID"].astype(str).str.upper()
                == str(selected_wsid).upper()
            ]

            tgl_str = f"{d:02d}/{month_str}/{year_str}"
            chart_tgl = f"{d:02d}-{month_str}"

            has_valid_day_data = False
            if not row_match.empty:
                r = row_match.iloc[0]
                for col_k in ["FRX", "DUR", "%", "FRX.1", "DUR.1", "%.1"]:
                    val_check = r.get(col_k)
                    if pd.notnull(val_check) and str(
                        val_check
                    ).strip() not in ["", "nan", "None"]:
                        has_valid_day_data = True
                        break

            if has_valid_day_data and not row_match.empty:
                r = row_match.iloc[0]
                hw_freq = (
                    int(float(r.get("FRX", 0)))
                    if pd.notnull(r.get("FRX"))
                    else 0
                )
                hw_dur = (
                    int(float(r.get("DUR", 0)))
                    if pd.notnull(r.get("DUR"))
                    else 0
                )
                hw_dt = (
                    round(float(r.get("%", 0)), 2)
                    if pd.notnull(r.get("%"))
                    else 0.0
                )

                p01_freq = (
                    int(float(r.get("FRX.1", 0)))
                    if pd.notnull(r.get("FRX.1"))
                    else 0
                )
                p01_dur = (
                    int(float(r.get("DUR.1", 0)))
                    if pd.notnull(r.get("DUR.1"))
                    else 0
                )
                p01_dt = (
                    round(float(r.get("%.1", 0)), 2)
                    if pd.notnull(r.get("%.1"))
                    else 0.0
                )

                p02_freq = (
                    int(float(r.get("FRX.2", 0)))
                    if pd.notnull(r.get("FRX.2"))
                    else 0
                )
                p02_dur = (
                    int(float(r.get("DUR.2", 0)))
                    if pd.notnull(r.get("DUR.2"))
                    else 0
                )
                p02_dt = (
                    round(float(r.get("%.2", 0)), 2)
                    if pd.notnull(r.get("%.2"))
                    else 0.0
                )

                p03_freq = (
                    int(float(r.get("FRX.3", 0)))
                    if pd.notnull(r.get("FRX.3"))
                    else 0
                )
                p03_dur = (
                    int(float(r.get("DUR.3", 0)))
                    if pd.notnull(r.get("DUR.3"))
                    else 0
                )
                p03_dt = (
                    round(float(r.get("%.3", 0)), 2)
                    if pd.notnull(r.get("%.3"))
                    else 0.0
                )

                p04_freq = (
                    int(float(r.get("FRX.4", 0)))
                    if pd.notnull(r.get("FRX.4"))
                    else 0
                )
                p04_dur = (
                    int(float(r.get("DUR.4", 0)))
                    if pd.notnull(r.get("DUR.4"))
                    else 0
                )
                p04_dt = (
                    round(float(r.get("%.4", 0)), 2)
                    if pd.notnull(r.get("%.4"))
                    else 0.0
                )

                uptime_val = round(100 - hw_dt, 2) if hw_dt <= 100 else 0.0
                chart_dates.append(chart_tgl)
                chart_uptime.append(uptime_val)

                row_bg = "#fcdad7" if (hw_freq > 0 or hw_dur > 0) else "#d5f5e3"
                dt_color = "#c0392b" if hw_dt > 0 else "#27ae60"

                row_html = f"""
                <tr style="background-color: {row_bg}; color: #000000;">
                    <td style="font-weight:bold; text-align:center;">{tgl_str}</td>
                    <td>{hw_freq}</td><td>{hw_dur}</td><td style="color:{dt_color}; font-weight:bold;">{hw_dt}</td>
                    <td>{p01_freq}</td><td>{p01_dur}</td><td>{p01_dt}</td>
                    <td>{p02_freq}</td><td>{p02_dur}</td><td>{p02_dt}</td>
                    <td>{p03_freq}</td><td>{p03_dur}</td><td>{p03_dt}</td>
                    <td>{p04_freq}</td><td>{p04_dur}</td><td>{p04_dt}</td>
                </tr>
                """
            else:
                row_html = f"""
                <tr style="background-color: #ffffff; color: #000000;">
                    <td style="font-weight:bold; text-align:center;">{tgl_str}</td>
                    <td></td><td></td><td></td>
                    <td></td><td></td><td></td>
                    <td></td><td></td><td></td>
                    <td></td><td></td><td></td>
                    <td></td><td></td><td></td>
                </tr>
                """

            table_frx_rows_html.append(row_html)

    table_frx_body = "".join(table_frx_rows_html)

    history_rows = []
    if not df_visit.empty:
        match_visit = pd.DataFrame()
        if "ID" in df_visit.columns:
            match_visit = df_visit[
                df_visit["ID"].astype(str).str.upper()
                == str(selected_wsid).upper()
            ]

        for _, vrow in match_visit.iterrows():
            spart_val = vrow.get("SPART", None)
            if pd.notnull(spart_val) and str(spart_val).strip() not in [
                "",
                "0",
                "nan",
                "None",
            ]:
                start_time = vrow.get("STARTED", "-")
                end_time = vrow.get("FINISHED", "-")
                history_rows.append(
                    {
                        "start": str(start_time),
                        "end": str(end_time),
                        "spart": str(spart_val).strip(),
                    }
                )

    history_html = (
        "".join(
            [
                f"<tr><td>{h['start']}</td><td>{h['end']}</td><td>{h['spart']}</td></tr>"
                for h in history_rows
            ]
        )
        if history_rows
        else '<tr><td colspan="3" style="text-align:center; color:#888; padding:15px;">Tidak ada history penggantian part</td></tr>'
    )

    html_code = f"""
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            * {{ box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; }}
            body {{ background-color: #f4f6f9; padding: 10px; color: #333; }}
            
            .top-banner {{ background-color: #ffffff; border-radius: 6px; padding: 12px 20px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 2px 4px rgba(0,0,0,0.08); margin-bottom: 15px; font-weight: 600; font-size: 15px; }}
            .dashboard-grid {{ display: grid; grid-template-columns: 28% 44% 28%; gap: 15px; }}
            .card {{ background: #ffffff; border-radius: 6px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.08); margin-bottom: 15px; }}
            .card-header {{ padding: 10px 15px; color: #ffffff; font-weight: bold; display: flex; justify-content: space-between; align-items: center; font-size: 14px; cursor: pointer; user-select: none; }}
            .card-header.lightblue {{ background-color: #3598db; }}
            .card-header.green {{ background-color: #27ae60; }}
            .card-header.purple {{ background-color: #9b59b6; }}
            .card-header.gold {{ background-color: #f1c40f; color: #000000; }}
            .header-engineer-info {{ color: #ffff00; font-size: 11px; font-weight: bold; margin-left: 8px; text-shadow: 0px 0px 2px rgba(0, 0, 0, 0.5); }}
            .info-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
            .info-table td {{ padding: 7px 12px; border-bottom: 1px solid #eef2f5; }}
            .info-table tr td:first-child {{ font-weight: bold; color: #444; width: 35%; }}
            .history-table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
            .history-table th {{ background: #f8f9fa; text-align: left; padding: 8px; border-bottom: 2px solid #dee2e6; }}
            .history-table td {{ padding: 8px; border-bottom: 1px solid #eee; }}
            
            .frx-excel-table {{ 
                width: 100%; 
                border-collapse: collapse; 
                font-size: 10px; 
                text-align: center; 
            }}
            .frx-excel-table th, .frx-excel-table td {{ 
                border: 1px solid #e2b667; 
                padding: 4px 2px;
            }}
            
            .p02-header-bg {{ background-color: #f39c12 !important; color: #000000 !important; font-size: 11px; font-weight: bold; }}
            .p02-sub-header {{ background-color: #fdf3e7 !important; color: #000000 !important; font-weight: bold; }}
            .table-scroll {{ max-height: 290px; overflow-y: auto; overflow-x: auto; position: relative; }}
        </style>
    </head>
    <body>
        <div class="top-banner">
            <div>WSID: <strong style="color: #2c3e50; font-size:17px;">{info_data['Wsid']}</strong> &nbsp;|&nbsp; <span style="background-color:#eef2f7; padding:4px 10px; border-radius:4px; font-weight:bold; color:#2c3e50;">📅 {selected_month}</span></div>
            <div>UPTIME BULANAN: {mtype}: <span style="color:{ut_color}; font-weight:bold;">{achieve_ut_str}</span> &nbsp;|&nbsp; STATUS: <span style="color:{status_color}; font-weight:bold;">{status_text}</span> &nbsp;|&nbsp; TOTAL FRX: <strong style="color:{frx_color}; font-size:16px;">{freq_dt_val}</strong></div>
        </div>

        <div class="dashboard-grid">
            <div class="card">
                <div class="card-header lightblue" onclick="toggleCard('infoContent', 'arrow1')"><span>Informasi Mesin</span><span id="arrow1">▼</span></div>
                <div id="infoContent"><table class="info-table">
                    <tr><td>Wsid</td><td>{info_data['Wsid']}</td></tr>
                    <tr><td>Ws Name</td><td>{info_data['Ws Name']}</td></tr>
                    <tr><td>Trx</td><td>{info_data['Trx']}</td></tr>
                    <tr><td>Serial No</td><td>{info_data['Serial No']}</td></tr>
                    <tr><td>Sn</td><td>{info_data['Sn']}</td></tr>
                    <tr><td>Address</td><td>{info_data['Address']}</td></tr>
                    <tr><td>City</td><td>{info_data['City']}</td></tr>
                    <tr><td>Province</td><td>{info_data['Province']}</td></tr>
                    <tr><td>Island</td><td>{info_data['Island']}</td></tr>
                    <tr><td>Location</td><td>{info_data['Location']}</td></tr>
                    <tr><td>Vip</td><td>{info_data['Vip']}</td></tr>
                    <tr><td>Mtype</td><td>{info_data['Mtype']}</td></tr>
                    <tr><td>Model</td><td>{info_data['Model']}</td></tr>
                    <tr><td>Sw Installed</td><td>{info_data['Sw Installed']}</td></tr>
                    <tr><td>Vendor</td><td>{info_data['Vendor']}</td></tr>
                </table></div>
            </div>

            <div>
                <div class="card">
                    <div class="card-header green" onclick="toggleCard('chartContent', 'arrow2')"><span>Grafik Uptime Harian</span><span id="arrow2">▼</span></div>
                    <div id="chartContent" style="padding:10px;"><canvas id="lineChart" height="100"></canvas></div>
                </div>

                <div class="card">
                    <div class="card-header purple" onclick="toggleCard('tableContent', 'arrow3')">
                        <div><span>Uptime Harian</span><span class="header-engineer-info">| {info_data['Wsid']} | {info_data['Engineer']}</span></div>
                        <span id="arrow3">▼</span>
                    </div>
                    <div id="tableContent" class="table-scroll">
                        <table class="frx-excel-table">
                            <thead>
                                <tr class="p02-header-bg">
                                    <th rowspan="2">Tgl</th><th colspan="3">HW TOTAL</th><th colspan="3">P01 - Pick Modul</th><th colspan="3">P02 - Presenter</th><th colspan="3">P03 - Reject</th><th colspan="3">P04 - Card Reader</th>
                                </tr>
                                <tr class="p02-sub-header">
                                    <th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th><th>Freq</th><th>Dur</th><th>DT</th>
                                </tr>
                            </thead>
                            <tbody>{table_frx_body}</tbody>
                        </table>
                    </div>
                </div>
            </div>

            <div class="card">
                <div class="card-header gold" onclick="toggleCard('historyContent', 'arrow4')"><span>History Spart</span><span id="arrow4">▼</span></div>
                <div id="historyContent"><table class="history-table">
                    <thead><tr><th>Start Time</th><th>End Time</th><th>SPart</th></tr></thead>
                    <tbody>{history_html}</tbody>
                </table></div>
            </div>
        </div>

        <script>
            function toggleCard(cId, aId) {{
                const c = document.getElementById(cId); const a = document.getElementById(aId);
                if (c.style.display === "none") {{ c.style.display = "block"; a.innerText = "▼"; }} else {{ c.style.display = "none"; a.innerText = "▲"; }}
            }}
            new Chart(document.getElementById('lineChart'), {{
                type: 'line',
                data: {{ labels: {json.dumps(chart_dates)}, datasets: [{{ label: 'Uptime (%)', data: {json.dumps(chart_uptime)}, borderColor: '#27ae60', backgroundColor: 'rgba(39, 174, 96, 0.1)', borderWidth: 2, fill: true }}] }},
                options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ min: 70, max: 100 }} }} }}
            }});
        </script>
    </body>
    </html>
    """
    st.components.v1.html(html_code, height=950, scrolling=True)


# =============================================================
# HALAMAN 2: UPTIME HARIAN (UPTIME DAILY CRM BCA)
# =============================================================
elif menu_option == "Uptime Harian":
    try:
        xls = pd.ExcelFile(target_excel_file)

        sheet_2_name = (
            xls.sheet_names[1]
            if len(xls.sheet_names) > 1
            else xls.sheet_names[0]
        )
        df_sheet2 = pd.read_excel(xls, sheet_name=sheet_2_name, header=None)

        ALL_MONTHS = [
            "JAN",
            "FEB",
            "MAR",
            "APR",
            "MEI",
            "JUN",
            "JUL",
            "AGUS",
            "SEPT",
            "OKT",
            "NOV",
            "DES",
        ]

        selected_m_num = int(selected_month.split("-")[1])
        start_idx = max(0, selected_m_num - 3)
        allowed_months = ALL_MONTHS[start_idx:selected_m_num]

        def build_uptime_daily_crm_bca_html(df, keep_months):
            html = """
            <style>
                .excel-wrapper {
                    display: flex;
                    flex-direction: column;
                    gap: 12px;
                }
                .excel-container {
                    max-height: 380px;
                    overflow-y: auto;
                    overflow-x: auto;
                    border: 1px solid #7f8c8d;
                    border-radius: 2px;
                    background-color: #ffffff;
                }
                .excel-table {
                    width: 100%;
                    border-collapse: collapse;
                    font-size: 11px;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                }
                .excel-table td, .excel-table th {
                    border: 1px solid #a6a6a6;
                    padding: 3px 5px;
                    white-space: nowrap;
                    text-align: center;
                }
                .header-row {
                    background-color: #d9e1f2 !important;
                    color: #000000 !important;
                    font-weight: bold;
                    position: sticky;
                    top: 0;
                    z-index: 2;
                }
                .col-no { width: 32px !important; text-align: center; }
                .col-cse { text-align: left !important; padding-left: 6px !important; font-weight: 500; }
                
                .row-section-title { 
                    background-color: #e2efda !important; 
                    color: #000000 !important;
                    font-weight: bold; 
                    text-align: left !important; 
                    padding: 5px 10px !important; 
                    font-size: 11px;
                }
                
                .bg-green { background-color: #c6efce !important; color: #006100 !important; font-weight: bold; }
                .bg-red { background-color: #ffc7ce !important; color: #9c0006 !important; font-weight: bold; }
                .row-summary { background-color: #f2f2f2 !important; font-weight: bold; }
            </style>
            <div class="excel-wrapper">
            """

            valid_col_indices = []
            header_row_idx = None

            for r_idx, r in df.iterrows():
                r_vals = [
                    "" if pd.isnull(x) else str(x).strip().upper()
                    for x in r.values
                ]
                if "NO" in r_vals and "CSE" in r_vals:
                    header_row_idx = r_idx
                    for col_i, col_v in enumerate(r_vals):
                        if col_i in [0, 1, 2]:
                            valid_col_indices.append(col_i)
                        elif col_v in ALL_MONTHS:
                            if col_v in keep_months:
                                valid_col_indices.append(col_i)
                    break

            if header_row_idx is not None:
                header_vals = [
                    "" if pd.isnull(x) else str(x).strip().upper()
                    for x in df.iloc[header_row_idx].values
                ]

                cse_start_row = header_row_idx + 1
                cse_end_row = len(df)
                for r_i in range(cse_start_row, len(df)):
                    r_txt = " ".join(
                        [
                            str(x).upper()
                            for x in df.iloc[r_i].values
                            if pd.notnull(x)
                        ]
                    )
                    if (
                        "TOTAL" in r_txt
                        or "PENGELOLA" in r_txt
                        or "BALI NUSRA" in r_txt
                    ):
                        cse_end_row = r_i
                        break

                for col_i in range(len(header_vals)):
                    if col_i in valid_col_indices:
                        continue

                    col_head_str = header_vals[col_i]

                    is_day_col = False
                    try:
                        f_v = float(col_head_str)
                        if 1 <= f_v <= 31:
                            is_day_col = True
                    except Exception:
                        pass

                    if is_day_col:
                        has_real_data = False
                        for r_idx in range(cse_start_row, cse_end_row):
                            cell_val = df.iloc[r_idx, col_i]
                            if pd.notnull(cell_val):
                                s_val = str(cell_val).strip().upper()
                                if s_val not in ["", "#DIV/0!", "NAN", "NONE"]:
                                    try:
                                        if float(s_val) > 0:
                                            has_real_data = True
                                            break
                                    except Exception:
                                        has_real_data = True
                                        break

                        if has_real_data:
                            valid_col_indices.append(col_i)

            rows_engineer = []
            rows_pengelola = []
            current_section = "ENGINEER"

            for row_idx, row in df.iterrows():
                row_vals = [
                    "" if pd.isnull(x) else str(x).strip() for x in row.values
                ]
                if not any(row_vals):
                    continue
                row_str = " ".join([v.upper() for v in row_vals])

                if row_idx == 0 and "UPTIME DAILY" in row_str:
                    title_text = [v for v in row_vals if v != ""][0]
                    st.markdown(
                        f"<h4 style='margin-bottom: 10px; color: #2c3e50;'>{title_text}</h4>",
                        unsafe_allow_html=True,
                    )
                    continue

                if "ENGINEER" in row_str or "UT :" in row_str:
                    continue

                if "PENGELOLA" in row_str:
                    current_section = "PENGELOLA"
                    continue

                if current_section == "ENGINEER":
                    rows_engineer.append((row_idx, row_vals, row_str))
                else:
                    rows_pengelola.append((row_idx, row_vals, row_str))

            def render_section_table(title, rows_list):
                sec_html = f'<div class="excel-container"><table class="excel-table"><tbody>'
                sec_html += f'<tr><td colspan="{len(valid_col_indices)}" class="row-section-title">{title}</td></tr>'

                for row_idx, row_vals, row_str in rows_list:
                    if "NO" in row_vals and "CSE" in row_vals:
                        sec_html += '<tr class="header-row">'
                        for col_idx in valid_col_indices:
                            v = (
                                row_vals[col_idx]
                                if col_idx < len(row_vals)
                                else ""
                            )
                            if v in ["#DIV/0!", "nan", "None"]:
                                v = ""
                            try:
                                num_v = float(v)
                                if num_v.is_integer():
                                    v = str(int(num_v))
                            except Exception:
                                pass

                            cls = (
                                "col-no"
                                if col_idx == 0
                                else ("col-cse" if col_idx == 1 else "")
                            )
                            sec_html += f'<td class="{cls}">{v}</td>'
                        sec_html += "</tr>"
                        continue

                    is_total = "TOTAL" in row_str
                    is_id_tidak_tercapai = "ID TIDAK TERCAPAI" in row_str

                    tr_class = (
                        ' class="row-summary"'
                        if (is_total or is_id_tidak_tercapai)
                        else ""
                    )
                    sec_html += f"<tr{tr_class}>"

                    for col_idx in valid_col_indices:
                        val = (
                            row_vals[col_idx]
                            if col_idx < len(row_vals)
                            else ""
                        )

                        if val == "" or val in ["#DIV/0!", "nan", "None"]:
                            sec_html += "<td></td>"
                            continue

                        cell_cls = []
                        if col_idx == 0:
                            cell_cls.append("col-no")
                        elif col_idx == 1:
                            cell_cls.append("col-cse")

                        display_val = val
                        try:
                            num_val = float(val)
                            if is_id_tidak_tercapai:
                                display_val = str(int(round(num_val)))
                                cell_cls.append("bg-red")
                            elif 0 < num_val <= 100 and col_idx >= 3:
                                display_val = f"{num_val:.2f}"
                                if num_val >= 99.20:
                                    cell_cls.append("bg-green")
                                else:
                                    cell_cls.append("bg-red")
                            elif num_val.is_integer():
                                display_val = str(int(num_val))
                            else:
                                display_val = f"{num_val:.2f}"
                        except Exception:
                            pass

                        cls_str = (
                            f' class="{" ".join(cell_cls)}"'
                            if cell_cls
                            else ""
                        )
                        sec_html += f"<td{cls_str}>{display_val}</td>"

                    sec_html += "</tr>"

                sec_html += "</tbody></table></div>"
                return sec_html

            html += render_section_table("ENGINEER", rows_engineer)

            if rows_pengelola:
                html += render_section_table("PENGELOLA", rows_pengelola)

            html += "</div>"
            return html

        table_html = build_uptime_daily_crm_bca_html(df_sheet2, allowed_months)
        st.components.v1.html(table_html, height=780, scrolling=True)

    except Exception as e:
        st.error(
            f"Terjadi kesalahan saat memproses data UPTIME DAILY CRM BCA: {e}"
        )


# =============================================================
# HALAMAN 3: UPTIME CSE BY TIPE MESIN
# =============================================================
elif menu_option == "Uptime CSE by Tipe Mesin":
    st.markdown(
        f"<h4 style='margin-bottom: 12px; color: #2c3e50; font-weight: 600;'>📊 Uptime CSE by Tipe Mesin ({selected_month})</h4>",
        unsafe_allow_html=True,
    )

    try:
        xls = pd.ExcelFile(target_excel_file)
        target_sheet = None
        for s in xls.sheet_names:
            if s.upper().strip() in [
                "BALI NUSRA",
                "UPTIME HARIAN",
                "UPTIME CSE",
            ]:
                target_sheet = s
                break

        if target_sheet:
            df_a15 = pd.read_excel(xls, target_sheet, skiprows=14, header=None)

            def build_exact_excel_html(df):
                html = """
                <style>
                    .excel-container {
                        max-height: 750px;
                        overflow-y: auto;
                        overflow-x: auto;
                        border: 1px solid #7f8c8d;
                        border-radius: 2px;
                        background-color: #ffffff;
                    }
                    .excel-table {
                        width: 100%;
                        border-collapse: collapse;
                        font-size: 11px;
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    }
                    .excel-table td, .excel-table th {
                        border: 1px solid #a6a6a6;
                        padding: 4px 6px;
                        white-space: nowrap;
                    }
                    
                    .col-no { width: 32px !important; text-align: center; }
                    .col-cse { width: 160px !important; text-align: left; padding-left: 6px !important; }
                    
                    .header-excel {
                        background-color: #8ea9db !important;
                        color: #000000 !important;
                        font-weight: bold;
                        text-align: center;
                        position: sticky;
                        top: 0;
                        z-index: 2;
                    }

                    .title-green-excel {
                        background-color: #c6efce !important;
                        color: #000000 !important;
                        font-weight: bold;
                        text-align: center;
                        font-size: 12px;
                    }

                    .row-total {
                        background-color: #d9e1f2 !important;
                        font-weight: bold;
                    }

                    .text-cse-red { color: #c00000 !important; font-weight: bold !important; }
                    .text-cse-black { color: #000000 !important; font-weight: normal !important; }
                    
                    .text-red { color: #c00000 !important; font-weight: bold; }
                    .text-green { color: #008000 !important; font-weight: bold; }
                    .bg-pink-red { background-color: #ffc7ce !important; color: #9c0006 !important; font-weight: bold; text-align: center; }
                    .bg-light-green { background-color: #c6efce !important; color: #006100 !important; font-weight: bold; text-align: center; }
                    .center { text-align: center; }
                </style>
                <div class="excel-container">
                <table class="excel-table">
                    <tbody>
                """

                for _, row in df.iterrows():
                    row_vals = [
                        "" if pd.isnull(x) else str(x).strip()
                        for x in row.values
                    ]

                    if not any(row_vals):
                        continue

                    row_str = " ".join([v.upper() for v in row_vals])

                    if "UPTIME" in row_str:
                        title_text = [v for v in row_vals if v != ""][0]
                        html += f'<tr><td colspan="9" class="title-green-excel">{title_text}</td></tr>'
                        continue

                    if "NO" in row_vals and (
                        "CSE" in row_vals or "PROVINSI" in row_vals
                    ):
                        html += '<tr class="header-excel">'
                        for col_idx, v in enumerate(row_vals[:9]):
                            cls = (
                                "col-no"
                                if col_idx == 0
                                else (
                                    "col-cse" if col_idx == 1 else "center"
                                )
                            )
                            html += f'<td class="{cls}">{v}</td>'
                        html += "</tr>"
                        continue

                    achieve_ut_val = None
                    try:
                        if len(row_vals) > 7 and row_vals[7] not in [
                            "",
                            "#DIV/0!",
                            "nan",
                            "None",
                        ]:
                            achieve_ut_val = float(row_vals[7])
                    except Exception:
                        pass

                    is_cse_red = (achieve_ut_val is not None) and (
                        achieve_ut_val < 99.20
                    )

                    is_total = "TOTAL" in row_str
                    tr_class = ' class="row-total"' if is_total else ""
                    html += f"<tr{tr_class}>"

                    for col_idx, val in enumerate(row_vals[:9]):
                        if val == "" or val in ["#DIV/0!", "nan", "None"]:
                            html += "<td></td>"
                            continue

                        cell_cls = []
                        if col_idx == 0:
                            cell_cls.append("col-no")
                        elif col_idx == 1:
                            cell_cls.append("col-cse")
                            if not is_total:
                                if is_cse_red:
                                    cell_cls.append("text-cse-red")
                                else:
                                    cell_cls.append("text-cse-black")
                        else:
                            cell_cls.append("center")

                        display_val = val
                        try:
                            num_val = float(val)
                            if col_idx == 7:
                                display_val = f"{num_val:.2f}"
                                cell_cls.append(
                                    "bg-light-green"
                                    if num_val >= 99.20
                                    else "bg-pink-red"
                                )
                            elif col_idx == 8:
                                display_val = f"{num_val:.2f}"
                                cell_cls.append(
                                    "text-red"
                                    if num_val > 0.80
                                    else "text-green"
                                )
                            elif num_val.is_integer():
                                display_val = str(int(num_val))
                            else:
                                display_val = f"{num_val:.2f}"
                        except Exception:
                            pass

                        cls_str = (
                            f' class="{" ".join(cell_cls)}"'
                            if cell_cls
                            else ""
                        )
                        html += f"<td{cls_str}>{display_val}</td>"

                    html += "</tr>"

                html += "</tbody></table></div>"
                return html

            table_html = build_exact_excel_html(df_a15)
            st.components.v1.html(table_html, height=750, scrolling=True)
        else:
            st.error(f"Sheet tidak ditemukan di {target_excel_file}!")
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")


# =============================================================
# HALAMAN 4: UPTIME PKT BY TIPE MESIN
# =============================================================
elif menu_option == "Uptime PKT by Tipe Mesin":
    st.markdown(
        f"<h4 style='margin-bottom: 12px; color: #2c3e50; font-weight: 600;'>📊 Uptime PKT by Tipe Mesin ({selected_month})</h4>",
        unsafe_allow_html=True,
    )

    try:
        xls = pd.ExcelFile(target_excel_file)

        sheet_4_name = "PKT TIPE MESIN"
        if sheet_4_name not in xls.sheet_names:
            if len(xls.sheet_names) >= 4:
                sheet_4_name = xls.sheet_names[3]

        df_pkt_raw = pd.read_excel(xls, sheet_name=sheet_4_name, header=None)

        def build_pkt_tipe_mesin_html(df):
            html = """
            <style>
                .excel-container {
                    max-height: 750px;
                    overflow-y: auto;
                    overflow-x: auto;
                    border: 1px solid #7f8c8d;
                    border-radius: 2px;
                    background-color: #ffffff;
                }
                .excel-table {
                    width: 100%;
                    border-collapse: collapse;
                    font-size: 11px;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                }
                .excel-table td, .excel-table th {
                    border: 1px solid #a6a6a6;
                    padding: 4px 6px;
                    white-space: nowrap;
                }
                
                .col-no { width: 32px !important; text-align: center; }
                .col-pkt { width: 180px !important; text-align: left; padding-left: 6px !important; }
                
                .header-excel {
                    background-color: #8ea9db !important;
                    color: #000000 !important;
                    font-weight: bold;
                    text-align: center;
                    position: sticky;
                    top: 0;
                    z-index: 2;
                }

                .title-green-excel {
                    background-color: #c6efce !important;
                    color: #000000 !important;
                    font-weight: bold;
                    text-align: center;
                    font-size: 12px;
                }

                .row-total {
                    background-color: #d9e1f2 !important;
                    font-weight: bold;
                }

                .text-pkt-red { color: #c00000 !important; font-weight: bold !important; }
                .text-pkt-black { color: #000000 !important; font-weight: normal !important; }
                
                .text-red { color: #c00000 !important; font-weight: bold; }
                .text-green { color: #008000 !important; font-weight: bold; }
                .bg-pink-red { background-color: #ffc7ce !important; color: #9c0006 !important; font-weight: bold; text-align: center; }
                .bg-light-green { background-color: #c6efce !important; color: #006100 !important; font-weight: bold; text-align: center; }
                .center { text-align: center; }
            </style>
            <div class="excel-container">
            <table class="excel-table">
                <tbody>
            """

            for _, row in df.iterrows():
                row_vals = [
                    "" if pd.isnull(x) else str(x).strip() for x in row.values
                ]

                if not any(row_vals):
                    continue

                row_str = " ".join([v.upper() for v in row_vals])

                if "UPTIME" in row_str:
                    title_text = [v for v in row_vals if v != ""][0]
                    html += f'<tr><td colspan="9" class="title-green-excel">{title_text}</td></tr>'
                    continue

                if "NO" in row_vals and ("PKT" in row_vals or "CSE" in row_vals):
                    html += '<tr class="header-excel">'
                    for col_idx, v in enumerate(row_vals[:9]):
                        cls = (
                            "col-no"
                            if col_idx == 0
                            else ("col-pkt" if col_idx == 1 else "center")
                        )
                        html += f'<td class="{cls}">{v}</td>'
                    html += "</tr>"
                    continue

                achieve_ut_val = None
                try:
                    if len(row_vals) > 7 and row_vals[7] not in [
                        "",
                        "#DIV/0!",
                        "nan",
                        "None",
                    ]:
                        achieve_ut_val = float(row_vals[7])
                except Exception:
                    pass

                is_pkt_red = (achieve_ut_val is not None) and (
                    achieve_ut_val < 99.20
                )

                is_total = "TOTAL" in row_str
                tr_class = ' class="row-total"' if is_total else ""
                html += f"<tr{tr_class}>"

                for col_idx, val in enumerate(row_vals[:9]):
                    if val == "" or val in ["#DIV/0!", "nan", "None"]:
                        html += "<td></td>"
                        continue

                    cell_cls = []
                    if col_idx == 0:
                        cell_cls.append("col-no")
                    elif col_idx == 1:
                        cell_cls.append("col-pkt")
                        if not is_total:
                            if is_pkt_red:
                                cell_cls.append("text-pkt-red")
                            else:
                                cell_cls.append("text-pkt-black")
                    else:
                        cell_cls.append("center")

                    display_val = val
                    try:
                        num_val = float(val)
                        if col_idx == 7:
                            display_val = f"{num_val:.2f}"
                            cell_cls.append(
                                "bg-light-green"
                                if num_val >= 99.20
                                else "bg-pink-red"
                            )
                        elif col_idx == 8:
                            display_val = f"{num_val:.2f}"
                            cell_cls.append(
                                "text-red"
                                if num_val > 0.80
                                else "text-green"
                            )
                        elif num_val.is_integer():
                            display_val = str(int(num_val))
                        else:
                            display_val = f"{num_val:.2f}"
                    except Exception:
                        pass

                    cls_str = (
                        f' class="{" ".join(cell_cls)}"' if cell_cls else ""
                    )
                    html += f"<td{cls_str}>{display_val}</td>"

                html += "</tr>"

            html += "</tbody></table></div>"
            return html

        table_html = build_pkt_tipe_mesin_html(df_pkt_raw)
        st.components.v1.html(table_html, height=750, scrolling=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses sheet PKT TIPE MESIN: {e}")


# =============================================================
# HALAMAN 5: RIWAYAT KUNJUNGAN (VISIT HISTORY)
# =============================================================
elif menu_option == "Riwayat Kunjungan (Visit)":
    st.markdown(
        f"<h4 style='margin-bottom: 12px; color: #2c3e50; font-weight: 600;'>📋 Riwayat Kunjungan WSID: <span style='color: #3498db;'>{selected_wsid}</span> ({selected_month})</h4>",
        unsafe_allow_html=True,
    )

    if df_visit.empty:
        st.warning("Data kunjungan (sheet 'MainVisit') tidak ditemukan pada file Master!")
    else:
        col_id_name = "ID" if "ID" in df_visit.columns else ("WSID" if "WSID" in df_visit.columns else None)

        if col_id_name:
            df_filtered_visit = df_visit[
                df_visit[col_id_name].astype(str).str.upper() == str(selected_wsid).upper()
            ]

            if df_filtered_visit.empty:
                st.info(f"Tidak ada histori kunjungan yang tercatat untuk WSID **{selected_wsid}**.")
            else:
                visit_rows_html = []
                for _, r in df_filtered_visit.iterrows():
                    ticket = str(r.get("TICKET", "-")) if pd.notnull(r.get("TICKET")) else "-"
                    svc_type = str(r.get("SVC_TYPE", "-")) if pd.notnull(r.get("SVC_TYPE")) else "-"
                    started = str(r.get("STARTED", "-")) if pd.notnull(r.get("STARTED")) else "-"
                    finished = str(r.get("FINISHED", "-")) if pd.notnull(r.get("FINISHED")) else "-"
                    solution = str(r.get("SOLUTION", "-")) if pd.notnull(r.get("SOLUTION")) else "-"

                    visit_rows_html.append(f"""
                    <tr>
                        <td style="font-weight:600; vertical-align: top;">{ticket}</td>
                        <td style="font-weight:600; vertical-align: top;">{svc_type}</td>
                        <td style="vertical-align: top;">{started}</td>
                        <td style="vertical-align: top;">{finished}</td>
                        <td style="text-align: left; line-height: 1.4; vertical-align: top;">{solution}</td>
                    </tr>
                    """)

                body_html = "".join(visit_rows_html)

                visit_table_html = f"""
                <!DOCTYPE html>
                <html lang="id">
                <head>
                    <style>
                        * {{ box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; }}
                        body {{ background-color: #f4f6f9; padding: 10px; }}
                        
                        .visit-card {{
                            background: #ffffff;
                            border-radius: 6px;
                            overflow: hidden;
                            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
                            border: 1px solid #dcdfe6;
                        }}
                        
                        .visit-header {{
                            background-color: #3598db;
                            color: #ffffff;
                            padding: 10px 15px;
                            font-weight: bold;
                            font-size: 14px;
                            display: flex;
                            justify-content: space-between;
                            align-items: center;
                        }}

                        .visit-table-container {{
                            max-height: 700px;
                            overflow-y: auto;
                            overflow-x: auto;
                        }}

                        .visit-table {{
                            width: 100%;
                            border-collapse: collapse;
                            font-size: 12px;
                            color: #333333;
                        }}

                        .visit-table th {{
                            background-color: #f8f9fa;
                            color: #2c3e50;
                            font-weight: bold;
                            text-align: left;
                            padding: 10px 12px;
                            border-bottom: 2px solid #dee2e6;
                            position: sticky;
                            top: 0;
                            z-index: 1;
                        }}

                        .visit-table td {{
                            padding: 12px;
                            border-bottom: 1px solid #eef2f5;
                        }}

                        .visit-table tr:hover {{
                            background-color: #f1f5f9;
                        }}
                    </style>
                </head>
                <body>
                    <div class="visit-card">
                        <div class="visit-header">
                            <span>Riwayat Kunjungan</span>
                            <span>▼</span>
                        </div>
                        <div class="visit-table-container">
                            <table class="visit-table">
                                <thead>
                                    <tr>
                                        <th style="width: 14%;">TICKET</th>
                                        <th style="width: 10%;">SVC_TYPE</th>
                                        <th style="width: 15%;">STARTED</th>
                                        <th style="width: 15%;">FINISHED</th>
                                        <th style="width: 46%;">SOLUTION</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {body_html}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </body>
                </html>
                """
                st.components.v1.html(visit_table_html, height=750, scrolling=True)
        else:
            st.dataframe(df_visit, use_container_width=True)
