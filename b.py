import time
import io
import json
import requests
import streamlit as st
from openpyxl import Workbook
from datetime import datetime

# ─── CONFIG ───────────────────────────────────────────────────────────────
BASE_URL = "https://api.subhxcosmo.in/api"
API_KEY = "DEMO6"
API_TYPE = "vehicle_num"
RETRIES = 2

# ─── ALL RESPONSE KEYS ─────────────────────────────────────────────────────
FULL_RESPONSE_KEYS = {
    "vehicle_number": "result.vnum",
    "owner": "owner",
    "mobile_no": "result.mobile_no",
    
    "challan_data": "result.challan_info.data",
    "challan_rc_info": "result.challan_info.rc_info",
    "challan_response_code": "result.challan_info.response_code",
    "challan_response_message": "result.challan_info.response_message",
    "challan_status": "result.challan_info.status",
    
    "vehicle_info_data": "result.vehicle_info.data",
    "vehicle_info_response_code": "result.vehicle_info.response_code",
    "vehicle_info_response_message": "result.vehicle_info.response_message",
    "vehicle_info_status": "result.vehicle_info.status",
    
    "debug_mob_raw": "result._debug.mob_raw",
    "debug_mob_error": "result._debug.mob_error",
    
    "success": "success",
    "cached": "cached",
    "proxy_used": "proxyUsed",
    "attempt": "attempt"
}

# ─── VALUE CONVERTER (to string) ───────────────────────────────────────────
def convert_value(value):
    """Convert any value to a string safe for Excel"""
    if value is None:
        return ""
    elif isinstance(value, bool):
        return "Yes" if value else "No"
    elif isinstance(value, (dict, list)):
        try:
            return json.dumps(value, ensure_ascii=False)
        except:
            return str(value)
    else:
        return str(value).strip()

# ─── NESTED VALUE FETCHER ──────────────────────────────────────────────────
def get_nested(data, path):
    try:
        keys = path.replace("]", "").split(".")
        for key in keys:
            if "[" in key:
                k, idx = key.split("[")
                data = data[k][int(idx)]
            else:
                data = data.get(key, {})
        return data if data != {} else ""
    except:
        return ""

# ─── FLATTEN RESPONSE (CLEAN) ──────────────────────────────────────────────
def flatten_full(data):
    flat = {}

    for key, path in FULL_RESPONSE_KEYS.items():
        value = get_nested(data, path)
        # Convert value to string to prevent Excel errors
        flat[key] = convert_value(value)

    return flat

# ─── FETCH API ─────────────────────────────────────────────────────────────
def fetch_vehicle(vehicle_no):
    url = f"{BASE_URL}?key={API_KEY}&type={API_TYPE}&term={vehicle_no}"

    for _ in range(RETRIES):
        try:
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                data = r.json()

                if not data.get("success"):
                    return {"status_flag": "FAILED"}

                flat = flatten_full(data)
                flat["status_flag"] = "FOUND" if flat.get("mobile_no") else "NOT_FOUND"

                return flat
        except:
            time.sleep(0.5)

    return {"status_flag": "ERROR"}

# ─── EXCEL BUILDER ─────────────────────────────────────────────────────────
def build_excel(rows):
    wb = Workbook()
    ws = wb.active

    headers = list(FULL_RESPONSE_KEYS.keys())
    ws.append(["Vehicle No", "Status"] + headers)

    for row in rows:
        values = [convert_value(row.get("vehicle_no")), convert_value(row.get("status_flag"))]
        for h in headers:
            values.append(convert_value(row.get(h, "")))
        ws.append(values)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── STREAMLIT UI ──────────────────────────────────────────────────────────
st.title("🚗 Vehicle Data Fetcher")

prefix = st.text_input("Vehicle Prefix", "RJ02CH")
start = st.number_input("Start Number", 1, 9999, 1)
end = st.number_input("End Number", 1, 9999, 10)
delay = st.slider("Delay (seconds)", 0.1, 2.0, 0.5)

if st.button("Start Fetching"):
    results = []

    for num in range(start, end + 1):
        vehicle_no = f"{prefix}{num:04d}"

        data = fetch_vehicle(vehicle_no)
        data["vehicle_no"] = vehicle_no

        results.append(data)

        st.write(f"{vehicle_no} → {data.get('status_flag')}")

        time.sleep(delay)

    excel = build_excel(results)

    st.download_button(
        "Download Excel",
        excel,
        file_name=f"vehicle_data_{datetime.now().strftime('%H%M%S')}.xlsx"
    )