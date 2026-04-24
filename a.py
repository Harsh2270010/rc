import time
import io
import requests
import streamlit as st
from openpyxl import Workbook
from datetime import datetime

# ─── CONFIG ───────────────────────────────────────────────────────────────
BASE_URL = "https://api.subhxcosmo.in/api"
API_KEY = "DEMO6"
API_TYPE = "vehicle"
RETRIES = 2

# ─── ALL RESPONSE KEYS ─────────────────────────────────────────────────────
FULL_RESPONSE_KEYS = {
    "registration_number": "result.registration_number",
    "vehicle_number": "result.chola2x_response.data.vehicleNumber",

    "owner_name": "result.chola2x_response.data.owner",
    "owner_father_name": "result.chola2x_response.data.ownerFatherName",
    "owner_count": "result.chola2x_response.data.ownerCount",

    "mobile_no": "result.chola2x_response.data.mobileNumber",
    "alternate_mobile": "result.chola2x_response.data.ownerMobileNumber",
    "contact_no": "result.vnum_chassis_response.mobile",

    "vehicle_type": "result.chola2x_response.vehicle_type",
    "vehicle_class": "result.chola2x_response.data.class",
    "vehicle_category": "result.chola2x_response.data.vehicleCategory",
    "model": "result.chola2x_response.data.model",
    "variant": "result.vnum_chassis_response.variant",
    "vehicle_name": "result.vnum_chassis_response.vehicle_name",
    "body_type": "result.chola2x_response.data.bodyType",

    "engine_no": "result.chola2x_response.data.engine",
    "chassis_no": "result.chola2x_response.data.chassis",
    "fuel_type": "result.chola2x_response.data.type",
    "cubic_capacity": "result.chola2x_response.data.vehicleCubicCapacity",
    "cylinders": "result.chola2x_response.data.vehicleCylindersNo",

    "manufacture_year": "result.vnum_chassis_response.manufacture_year",
    "manufacture_month": "result.vnum_chassis_response.manufacture_month",
    "manufacturer": "result.chola2x_response.data.vehicleManufacturerName",

    "registration_date": "result.chola2x_response.data.regDate",
    "rc_expiry_date": "result.chola2x_response.data.rcExpiryDate",
    "status": "result.chola2x_response.data.status",

    "rto_code": "result.chola2x_response.data.rtoCode",
    "rto_name": "result.vnum_chassis_response.rto_name",
    "rto_location": "result.chola2x_response.data.regAuthority",

    "insurance_upto": "result.chola2x_response.data.vehicleInsuranceUpto",
    "insurance_company": "result.chola2x_response.data.vehicleInsuranceCompanyName",
    "insurance_policy_no": "result.chola2x_response.data.vehicleInsurancePolicyNumber",

    "pucc_upto": "result.chola2x_response.data.puccUpto",

    "seating_capacity": "result.chola2x_response.data.vehicleSeatCapacity",
    "gross_weight": "result.chola2x_response.data.grossVehicleWeight",
    "unladen_weight": "result.chola2x_response.data.unladenWeight",

    "present_address": "result.chola2x_response.data.presentAddress",
    "district": "result.chola2x_response.data.splitPresentAddress.district[0]",
    "state": "result.chola2x_response.data.splitPresentAddress.state[0][0]",
    "pincode": "result.chola2x_response.data.splitPresentAddress.pincode",

    "vehicle_tax_upto": "result.chola2x_response.data.vehicleTaxUpto",
    "financer": "result.chola2x_response.data.rcFinancer",

    "is_commercial": "result.chola2x_response.data.isCommercial",
    "blacklist_status": "result.chola2x_response.data.blacklistStatus",

    "source": "result.chola2x_response.source",
    "cached": "cached",
    "proxy_used": "proxyUsed",
    "attempt": "attempt"
}

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

        # Mobile fallback (optional)
        if key == "mobile_no" and not value:
            value = (
                get_nested(data, "result.chola2x_response.data.mobile") or
                get_nested(data, "result.chola2x_response.data.phone") or
                get_nested(data, "result.vnum_chassis_response.contact") or
                ""
            )

        flat[key] = value if value else ""

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
                flat["status_flag"] = "FOUND" if flat.get("owner_name") else "NOT_FOUND"

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
        ws.append(
            [row.get("vehicle_no"), row.get("status_flag")]
            + [row.get(h, "") for h in headers]
        )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── STREAMLIT UI ──────────────────────────────────────────────────────────
st.title("🚗 Vehicle Data Fetcher")

prefix = st.text_input("Vehicle Prefix", "RJ60CC")
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