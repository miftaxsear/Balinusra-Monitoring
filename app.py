import json
import os
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Dashboard FRX DT", layout="wide")


# 2. Fungsi Membaca Master Excel & Seluruh Sheet Harian
@st.cache_data
def load_data(file_path):
    xls = pd.ExcelFile(file_path)

    # Sheet FRX DT (Utama / Fallback)
    df_frx_raw = (
        pd.read_excel(xls, "FRX DT")
        if "FRX DT" in xls.sheet_names
        else pd.DataFrame()
    )

    # Sheet SITE BLNS untuk Informasi Mesin
    df_site = (
        pd.read_excel(xls, "SITE BLNS")
        if "SITE BLNS" in xls.sheet_names
        else pd.DataFrame()
    )

    # Sheet MainVisit untuk History Spart
    df_visit = (
        pd.read_excel(xls, "MainVisit")
        if "MainVisit" in xls.sheet_names
        else pd.DataFrame()
    )

    # Sheet atm total untuk Uptime Bulanan & Total FRX
    df_atm_summary = pd.DataFrame()
    if "atm total" in xls.sheet_names:
        try:
            df_atm_summary = pd.read_excel(xls, "atm total", skiprows=5)
        except:
            pass

    # Membaca Sheet Harian '1', '2', ..., '31'
    daily_sheets = {}
    num_sheets = [s for s in xls.sheet_names if s.isdigit()]
    num_sheets.sort(key=lambda x: int(x))

    for sname in num_sheets:
        try:
            df_day = pd.read_excel(xls, sname, header=5)
            if "WSID" in df_day.columns:
                daily_sheets[int(sname)] = df_day
        except:
            pass

    return df_frx_raw, df_site, df_visit, df_atm_summary, daily_sheets


excel_file = "MASTER.xlsx"

if os.path.exists(excel_file):
    df_frx_raw, df_site, df_visit, df_atm_summary, daily_sheets = load_data(
        excel_file
    )
else:
    st.error(f"File '{excel_file}' tidak ditemukan di folder kerja!")
    st.stop()

# 3. Ekstraksi Daftar WSID untuk Input Search Box
list_wsid = []
if not df_site.empty and "ID" in df_site.columns:
    list_wsid = df_site["ID"].dropna().unique().tolist()
elif not df_atm_summary.empty and "WSID" in df_atm_summary.columns:
    list_wsid = df_atm_summary["WSID"].dropna().unique().tolist()
elif not df_visit.empty and "ID" in df_visit.columns:
    list_wsid = df_visit["ID"].dropna().unique().tolist()

if "ZRK5" not in list_wsid:
    list_wsid.insert(0, "ZRK5")

# Sidebar untuk Input Search WSID
st.sidebar.header("🔍 Filter Dashboard")
selected_wsid = st.sidebar.selectbox(
    "Pilih / Ketik WSID:", options=list_wsid, index=0
)

# 4. Informasi Mesin (SITE BLNS)
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

# 5. Nilai Uptime Bulanan & Total FRX
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
            except:
                pass

achieve_ut_str = f"{round(achieve_ut_float, 2)}%"

# 6. Logika Status "Tercapai / Tidak tercapai" & Warna
mtype = info_data["Mtype"]
target_ut = 99.75 if "ATM" in mtype else 99.20
is_tercapai = achieve_ut_float >= target_ut

status_text = "Tercapai" if is_tercapai else "Tidak tercapai"
status_color = "#27ae60" if is_tercapai else "#e74c3c"
ut_color = status_color
frx_color = "#e74c3c" if freq_dt_val > 0 else "#27ae60"

# 7. Memproses Data Sheet FRX DT untuk WSID Terpilih (Otomatis dari Sheet Harian)
chart_dates = []
chart_uptime = []
table_frx_rows_html = []

if daily_sheets:
    sorted_days = sorted(daily_sheets.keys())
    for d in sorted_days:
        df_d = daily_sheets[d]
        row_match = df_d[
            df_d["WSID"].astype(str).str.upper() == str(selected_wsid).upper()
        ]

        tgl_str = f"{d:02d}/09/2026"
        chart_tgl = f"{d:02d}-09"
        chart_dates.append(chart_tgl)

        if not row_match.empty:
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
        else:
            hw_freq = hw_dur = hw_dt = 0
            p01_freq = p01_dur = p01_dt = 0
            p02_freq = p02_dur = p02_dt = 0
            p03_freq = p03_dur = p03_dt = 0
            p04_freq = p04_dur = p04_dt = 0

        uptime_val = round(100 - hw_dt, 2) if hw_dt <= 100 else 0.0
        chart_uptime.append(uptime_val)

        # Skema Warna Baris (Merah Muda jika ada DT/Freq, Hijau Muda jika normal)
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
        table_frx_rows_html.append(row_html)

table_frx_body = "".join(table_frx_rows_html)

# 8. Ekstraksi Data History Spart dari Sheet MainVisit
history_rows = []
if not df_visit.empty:
    match_visit = pd.DataFrame()
    if "ID" in df_visit.columns:
        match_visit = df_visit[
            df_visit["ID"].astype(str).str.upper() == str(selected_wsid).upper()
        ]

    if match_visit.empty and "WSID" in df_visit.columns:
        match_visit = df_visit[
            df_visit["WSID"]
            .astype(str)
            .str.upper()
            .str.contains(str(selected_wsid).upper())
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

            if pd.notnull(start_time):
                try:
                    start_time = pd.to_datetime(start_time).strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                except:
                    start_time = str(start_time)

            if pd.notnull(end_time):
                try:
                    end_time = pd.to_datetime(end_time).strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                except:
                    end_time = str(end_time)

            history_rows.append(
                {
                    "start": start_time,
                    "end": end_time,
                    "spart": str(spart_val).strip(),
                }
            )

if history_rows:
    history_html = "".join(
        [
            f"<tr><td>{h['start']}</td><td>{h['end']}</td><td>{h['spart']}</td></tr>"
            for h in history_rows
        ]
    )
else:
    history_html = '<tr><td colspan="3" style="text-align:center; color:#888; padding:15px;">Tidak ada history penggantian part</td></tr>'

# 9. Render Dashboard HTML
html_code = f"""
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{ box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; }}
        body {{ background-color: #f4f6f9; padding: 10px; color: #333; }}
        
        .top-banner {{
            background-color: #ffffff;
            border-radius: 6px;
            padding: 12px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
            margin-bottom: 15px;
            font-weight: 600;
            font-size: 15px;
        }}

        .dashboard-grid {{
            display: grid;
            grid-template-columns: 28% 44% 28%;
            gap: 15px;
        }}

        .card {{ background: #ffffff; border-radius: 6px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.08); margin-bottom: 15px; }}
        .card-header {{ padding: 10px 15px; color: #ffffff; font-weight: bold; display: flex; justify-content: space-between; font-size: 14px; }}
        
        .card-header.lightblue {{ background-color: #3598db; }}
        .card-header.green {{ background-color: #27ae60; }}
        .card-header.purple {{ background-color: #9b59b6; }}
        .card-header.gold {{ background-color: #f1c40f; color: #000000; }}

        .info-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        .info-table td {{ padding: 7px 12px; border-bottom: 1px solid #eef2f5; }}
        .info-table tr td:first-child {{ font-weight: bold; color: #444; width: 35%; }}

        .chart-container {{ padding: 10px; }}

        .history-table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
        .history-table th {{ background: #f8f9fa; text-align: left; padding: 8px; border-bottom: 2px solid #dee2e6; }}
        .history-table td {{ padding: 8px; border-bottom: 1px solid #eee; }}

        /* Style Tabel Rincian FRX DT Excel */
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
        
        /* Header Kolom Menggunakan Warna Hitam (#000000) untuk Font */
        .p02-header-bg {{
            background-color: #f39c12 !important;
            color: #000000 !important;
            font-size: 11px;
            font-weight: bold;
        }}
        .p02-sub-header {{
            background-color: #fdf3e7 !important;
            color: #000000 !important;
            font-weight: bold;
        }}
        
        .table-scroll {{
            max-height: 290px;
            overflow-y: auto;
            overflow-x: auto;
        }}
    </style>
</head>
<body>

    <!-- Header Top Banner -->
    <div class="top-banner">
        <div>
            <span>WSID: <strong style="color: #2c3e50; font-size:17px;">{info_data['Wsid']}</strong></span> &nbsp;|&nbsp; 
            <select><option>2026-09</option></select>
        </div>
        <div>
            <span>UPTIME BULANAN: {mtype}: <span style="color:{ut_color}; font-weight:bold;">{achieve_ut_str}</span></span> &nbsp;|&nbsp; 
            <span>STATUS: <span style="color:{status_color}; font-weight:bold;">{status_text}</span></span> &nbsp;|&nbsp; 
            <span>TOTAL FRX: <strong style="color:{frx_color}; font-size:16px;">{freq_dt_val}</strong></span>
        </div>
    </div>

    <!-- Grid Layout -->
    <div class="dashboard-grid">
        
        <!-- Panel 1: Informasi Mesin -->
        <div class="card">
            <div class="card-header lightblue">
                Informasi Mesin <span>▼</span>
            </div>
            <table class="info-table">
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
            </table>
        </div>

        <!-- Panel 2: Grafik Uptime & Tabel Uptime Harian -->
        <div>
            <!-- Grafik Line Uptime Harian -->
            <div class="card">
                <div class="card-header green">
                    Grafik Uptime Harian <span>▼</span>
                </div>
                <div class="chart-container">
                    <canvas id="lineChart" height="100"></canvas>
                </div>
            </div>

            <!-- Tabel Uptime Harian (Font Header Hitam) -->
            <div class="card">
                <div class="card-header purple">
                    Uptime Harian <span>▼</span>
                </div>
                <div class="table-scroll">
                    <table class="frx-excel-table">
                        <thead>
                            <tr class="p02-header-bg">
                                <th rowspan="2" style="color: #000000;">Tgl</th>
                                <th colspan="3" style="color: #000000;">HW TOTAL</th>
                                <th colspan="3" style="color: #000000;">P01 - Pick Modul</th>
                                <th colspan="3" style="color: #000000;">P02 - Presenter</th>
                                <th colspan="3" style="color: #000000;">P03 - Reject</th>
                                <th colspan="3" style="color: #000000;">P04 - Card Reader</th>
                            </tr>
                            <tr class="p02-sub-header">
                                <th style="color: #000000;">Freq</th><th style="color: #000000;">Dur</th><th style="color: #000000;">DT</th>
                                <th style="color: #000000;">Freq</th><th style="color: #000000;">Dur</th><th style="color: #000000;">DT</th>
                                <th style="color: #000000;">Freq</th><th style="color: #000000;">Dur</th><th style="color: #000000;">DT</th>
                                <th style="color: #000000;">Freq</th><th style="color: #000000;">Dur</th><th style="color: #000000;">DT</th>
                                <th style="color: #000000;">Freq</th><th style="color: #000000;">Dur</th><th style="color: #000000;">DT</th>
                            </tr>
                        </thead>
                        <tbody>
                            {table_frx_body}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Panel 3: History Spart -->
        <div class="card">
            <div class="card-header gold">
                History Spart <span>▼</span>
            </div>
            <table class="history-table">
                <thead>
                    <tr>
                        <th>Start Time</th>
                        <th>End Time</th>
                        <th>SPart</th>
                    </tr>
                </thead>
                <tbody>
                    {history_html}
                </tbody>
            </table>
        </div>

    </div>

    <script>
        const dates = {json.dumps(chart_dates)};
        const uptimeData = {json.dumps(chart_uptime)};
        
        new Chart(document.getElementById('lineChart'), {{
            type: 'line',
            data: {{
                labels: dates,
                datasets: [{{
                    label: 'Uptime (%)',
                    data: uptimeData,
                    borderColor: '#27ae60',
                    backgroundColor: 'rgba(39, 174, 96, 0.1)',
                    borderWidth: 2,
                    fill: true,
                    tension: 0.1
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ min: 70, max: 100 }} }} }}
        }});
    </script>
</body>
</html>
"""

st.components.v1.html(html_code, height=920, scrolling=True)