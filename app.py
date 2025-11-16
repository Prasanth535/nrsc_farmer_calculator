from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    send_file,
    jsonify,
)
from datetime import datetime, timedelta
import io

from docx import Document
from reportlab.pdfgen import canvas

import pandas as pd

app = Flask(__name__)
app.secret_key = "CHANGE_THIS_SECRET_KEY"  # change in production

# ================== FILE PATHS ==================

CROP_EXCEL_PATH = "crop_variety_stcr.xlsx"   # crops, varieties, season, duration, Rec_N, Rec_P2O5, Rec_K2O
STCR_CSV_PATH = "stcr_formulas.csv"         # crop-wise STCR formulas


# ================== LOAD CROP–VARIETY DATA FROM EXCEL ==================

try:
    crop_df = pd.read_excel(CROP_EXCEL_PATH)
except Exception as e:
    print("Error loading Excel file:", e)
    crop_df = pd.DataFrame(
        columns=["Crop", "Variety", "Season", "Duration_days", "Rec_N", "Rec_P2O5", "Rec_K2O"]
    )

for col in ["Crop", "Variety", "Season"]:
    if col in crop_df.columns:
        crop_df[col] = crop_df[col].astype(str).str.strip()

if "Crop" in crop_df.columns:
    CROPS = sorted(crop_df["Crop"].dropna().unique().tolist())
else:
    CROPS = []

CROP_VARIETY_MAP = {}
if "Crop" in crop_df.columns and "Variety" in crop_df.columns:
    for crop_name in CROPS:
        vars_list = (
            crop_df.loc[crop_df["Crop"] == crop_name, "Variety"]
            .dropna()
            .unique()
            .tolist()
        )
        CROP_VARIETY_MAP[crop_name] = sorted(vars_list)


def get_crop_row(crop_name, variety_name):
    """Return row (dict) from Excel for (crop, variety) or None."""
    if crop_df.empty:
        return None
    crop_name = str(crop_name).strip()
    variety_name = str(variety_name).strip()
    sub = crop_df[
        (crop_df["Crop"] == crop_name)
        & (crop_df["Variety"] == variety_name)
    ]
    if sub.empty:
        # If no exact variety match, try crop only (first row)
        sub = crop_df[crop_df["Crop"] == crop_name]
        if sub.empty:
            return None
    return sub.iloc[0].to_dict()


# ================== LOAD STCR FORMULAS FROM CSV ==================

try:
    stcr_df = pd.read_csv(STCR_CSV_PATH)
except Exception as e:
    print("Error loading STCR CSV file:", e)
    stcr_df = pd.DataFrame(
        columns=[
            "Crop",
            "Variety",
            "TargetYield_t_ha",
            "Formula_N",
            "Formula_P2O5",
            "Formula_K2O",
        ]
    )

for col in ["Crop", "Variety"]:
    if col in stcr_df.columns:
        stcr_df[col] = stcr_df[col].astype(str).str.strip()


def get_stcr_row(crop_name, variety_name):
    """
    Return STCR row dict for given crop & variety.
    Priority:
    1. Exact (Crop, Variety)
    2. (Crop, Variety == 'ALL' / 'All' / '' / NaN)
    """
    if stcr_df.empty:
        return None

    crop_name = str(crop_name).strip()
    variety_name = str(variety_name).strip()

    exact = stcr_df[
        (stcr_df["Crop"] == crop_name)
        & (stcr_df["Variety"] == variety_name)
    ]
    if not exact.empty:
        return exact.iloc[0].to_dict()

    fallback = stcr_df[
        (stcr_df["Crop"] == crop_name)
        & (stcr_df["Variety"].str.upper().isin(["ALL", "", "NAN"]))
    ]
    if not fallback.empty:
        return fallback.iloc[0].to_dict()

    return None


def safe_eval_formula(formula_str, context):
    """
    Safely evaluate STCR formula string with given context.
    Allowed names: T, SN, SP, SK, max, min.
    """
    if not isinstance(formula_str, str) or not formula_str.strip():
        return 0.0

    allowed_names = {
        "T": context.get("T", 0.0),
        "SN": context.get("SN", 0.0),
        "SP": context.get("SP", 0.0),
        "SK": context.get("SK", 0.0),
        "max": max,
        "min": min,
    }
    try:
        value = eval(formula_str, {"__builtins__": {}}, allowed_names)
        return float(value)
    except Exception as e:
        print("Error evaluating formula:", formula_str, "Error:", e)
        return 0.0


# ================== TELANGANA SOIL TYPES ==================

TELANGANA_SOIL_TYPES = [
    "Red Clayey",
    "Red Loamy",
    "Red Shallow Loamy",
    "Red Shallow Clayey",
    "Red Shallow Gravelly Loam",
    "Red Shallow Gravelly Clayey",
    "Red Gravelly Loam",
    "Red Gravelly Clayey",
    "Red Calcareous Clayey",
    "Red Calcareous Gravelly Clayey",
    "Red Shallow Calcareous Gravelly Loam",
    "Medium Calcareous Black",
    "Deep Calcareous Black",
    "Deep Black",
    "Shallow Black",
    "Black (General)",
    "Red Shallow Gravelly Clay",
    "Red Shallow Gravelly Loamy",
    "Red Shallow Gravelly",
    "Red Gravelly",
    "Lateritic Gravelly Clayey",
    "Red Gravelly Clayey Loamy",
    "Alluvio Colluvial Clay",
    "Alluvio Colluvial Loamy",
    "Alluvio Colluvial Clayey",
    "Alluvio–Colluvial Clay",
    "Alluvio Colluvial Clay Loamy",
    "Alluvial Soil",
    "Alluvial Colluvial Loamy",
    "Alluvial Colluvial Clayey",
    "Saline Sodic Soil",
    "Calcareous Black",
    "Brown Forest Soils",
    "Rock Lands",
    "Water Body",
]


# ================== FERTILIZER SPLIT SCHEDULE ==================

FERT_SPLITS = {
    "Rice": [
        {"stage": "Basal",              "das_min": 0,  "das_max": 7,   "N_pct": 50, "P_pct": 100, "K_pct": 50},
        {"stage": "Tillering",          "das_min": 20, "das_max": 25,  "N_pct": 25, "P_pct": 0,   "K_pct": 25},
        {"stage": "Panicle initiation", "das_min": 40, "das_max": 45,  "N_pct": 25, "P_pct": 0,   "K_pct": 25},
    ],
    "Maize": [
        {"stage": "Basal",      "das_min": 0,  "das_max": 0,   "N_pct": 50, "P_pct": 100, "K_pct": 50},
        {"stage": "Knee-high",  "das_min": 25, "das_max": 30,  "N_pct": 25, "P_pct": 0,   "K_pct": 25},
        {"stage": "Tasseling",  "das_min": 45, "das_max": 50,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
    ],
    "Sorghum": [
        {"stage": "Basal",  "das_min": 0,  "das_max": 0,   "N_pct": 50, "P_pct": 100, "K_pct": 50},
        {"stage": "30 DAS", "das_min": 30, "das_max": 30,  "N_pct": 50, "P_pct": 0,   "K_pct": 50},
    ],
    "Wheat": [
        {"stage": "Basal",        "das_min": 0,  "das_max": 0,   "N_pct": 50, "P_pct": 100, "K_pct": 100},
        {"stage": "CRI",          "das_min": 20, "das_max": 22,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
        {"stage": "45–50 DAS",    "das_min": 45, "das_max": 50,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
    ],
    "Finger millet": [
        {"stage": "Basal",  "das_min": 0,  "das_max": 0,   "N_pct": 50, "P_pct": 100, "K_pct": 100},
        {"stage": "30 DAS", "das_min": 30, "das_max": 30,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
        {"stage": "45 DAS", "das_min": 45, "das_max": 45,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
    ],
    "Red gram": [
        {"stage": "Basal", "das_min": 0, "das_max": 0, "N_pct": 100, "P_pct": 100, "K_pct": 100},
    ],
    "Green gram": [
        {"stage": "Basal", "das_min": 0, "das_max": 0, "N_pct": 100, "P_pct": 100, "K_pct": 100},
    ],
    "Black gram": [
        {"stage": "Basal", "das_min": 0, "das_max": 0, "N_pct": 100, "P_pct": 100, "K_pct": 100},
    ],
    "Groundnut": [
        {"stage": "Basal", "das_min": 0, "das_max": 0, "N_pct": 100, "P_pct": 100, "K_pct": 100},
    ],
    "Soybean": [
        {"stage": "Basal", "das_min": 0, "das_max": 0, "N_pct": 100, "P_pct": 100, "K_pct": 100},
    ],
    "Cotton": [
        {"stage": "Basal",  "das_min": 0,  "das_max": 0,   "N_pct": 25, "P_pct": 100, "K_pct": 50},
        {"stage": "30 DAS", "das_min": 30, "das_max": 30,  "N_pct": 25, "P_pct": 0,   "K_pct": 25},
        {"stage": "60 DAS", "das_min": 60, "das_max": 60,  "N_pct": 25, "P_pct": 0,   "K_pct": 25},
        {"stage": "90 DAS", "das_min": 90, "das_max": 90,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
    ],
    "Chilli": [
        {"stage": "Basal",       "das_min": 0,  "das_max": 0,   "N_pct": 30, "P_pct": 100, "K_pct": 50},
        {"stage": "30 DAS",      "das_min": 30, "das_max": 30,  "N_pct": 25, "P_pct": 0,   "K_pct": 25},
        {"stage": "Flowering",   "das_min": 45, "das_max": 50,  "N_pct": 25, "P_pct": 0,   "K_pct": 25},
        {"stage": "Fruit set",   "das_min": 60, "das_max": 70,  "N_pct": 20, "P_pct": 0,   "K_pct": 0},
    ],
    "Turmeric": [
        {"stage": "Basal",  "das_min": 0,  "das_max": 0,   "N_pct": 50, "P_pct": 100, "K_pct": 100},
        {"stage": "45 DAS", "das_min": 45, "das_max": 45,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
        {"stage": "90 DAS", "das_min": 90, "das_max": 90,  "N_pct": 25, "P_pct": 0,   "K_pct": 0},
    ],
}


# ================== HELPER FUNCTIONS ==================

def compute_averages(readings):
    if not readings:
        return {}
    keys = readings[0].keys()
    averages = {}
    for k in keys:
        try:
            averages[k] = sum(float(r.get(k, 0.0) or 0.0) for r in readings) / len(readings)
        except Exception:
            averages[k] = 0.0
    return averages


def generate_irrigation_schedule(sowing_date_str, duration_days):
    if not sowing_date_str or not duration_days:
        return []
    try:
        sowing_date = datetime.strptime(sowing_date_str, "%Y-%m-%d").date()
    except ValueError:
        return []
    schedule = []
    d = sowing_date + timedelta(days=7)
    while (d - sowing_date).days <= duration_days:
        schedule.append(d.isoformat())
        d += timedelta(days=7)
    return schedule


def soil_type_factor(soil_type):
    if not soil_type:
        return 1.0
    st = soil_type.lower()
    if "deep black" in st or "calcareous" in st:
        return 0.9
    if "saline" in st or "sodic" in st:
        return 0.85
    if "shallow" in st or "gravelly" in st:
        return 1.0
    return 1.0


def npk_to_fertilizer_form(n_kg, p2o5_kg, k2o_kg):
    dap_needed = p2o5_kg / 0.46 if p2o5_kg else 0.0
    n_from_dap = dap_needed * 0.18
    n_remaining = max(0.0, n_kg - n_from_dap)
    urea_needed = n_remaining / 0.46 if n_remaining else 0.0
    mop_needed = k2o_kg / 0.60 if k2o_kg else 0.0
    return (
        f"Urea: {urea_needed:.1f} kg/ha, "
        f"DAP: {dap_needed:.1f} kg/ha, "
        f"MOP: {mop_needed:.1f} kg/ha"
    )


def generate_fertilizer_schedule(crop_name, sowing_date_str, n_total, p_total, k_total):
    if not sowing_date_str:
        return []
    try:
        sowing_date = datetime.strptime(sowing_date_str, "%Y-%m-%d").date()
    except ValueError:
        return []
    splits = FERT_SPLITS.get(crop_name, [])
    if not splits:
        splits = [
            {"stage": "Basal", "das_min": 0, "das_max": 0, "N_pct": 100, "P_pct": 100, "K_pct": 100}
        ]
    schedule = []
    for s in splits:
        das_min = s.get("das_min", 0) or 0
        das_apply = 0 if das_min == 0 else das_min
        app_date = sowing_date + timedelta(days=das_apply)

        N_pct = s.get("N_pct", 0) or 0
        P_pct = s.get("P_pct", 0) or 0
        K_pct = s.get("K_pct", 0) or 0

        n_amt = round(n_total * N_pct / 100.0, 2) if N_pct > 0 else 0.0
        p_amt = round(p_total * P_pct / 100.0, 2) if P_pct > 0 else 0.0
        k_amt = round(k_total * K_pct / 100.0, 2) if K_pct > 0 else 0.0

        schedule.append(
            {
                "stage": s.get("stage", ""),
                "das": das_apply,
                "date": app_date.isoformat(),
                "N": n_amt,
                "P2O5": p_amt,
                "K2O": k_amt,
            }
        )
    return schedule


# ================== ROUTES ==================

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        lang = request.form.get("language", "en")
        session["language"] = lang
        return redirect(url_for("farmer_details"))
    return render_template("index.html")


@app.route("/farmer-details", methods=["GET", "POST"])
def farmer_details():
    if request.method == "POST":
        farmer = {
            "name": request.form.get("name", "").strip(),
            "mobile": request.form.get("mobile", "").strip(),
            "village": request.form.get("village", "").strip(),
            "mandal": request.form.get("mandal", "").strip(),
            "district": request.form.get("district", "").strip(),
        }
        session["farmer"] = farmer
        return redirect(url_for("crop_details"))

    farmer = session.get("farmer", {})
    return render_template("farmer_details.html", farmer=farmer)


@app.route("/crop-details", methods=["GET", "POST"])
def crop_details():
    """
    Crop & soil details page.
    Frontend JS:
    - gets GPS location via browser
    - filters varieties by crop
    - fetches season/duration/NPK via /api/crop-info
    """
    crop_session = session.get("crop", {})

    if request.method == "POST":
        selected_crop = request.form.get("crop") or ""
        selected_variety = request.form.get("variety") or ""
        soil_type = request.form.get("soil_type") or ""
        gps_location = request.form.get("gps_location", "").strip()
        sowing_date = request.form.get("sowing_date", "").strip()

        ty_str = request.form.get("target_yield", "").strip()
        try:
            target_yield = float(ty_str) if ty_str else None
        except ValueError:
            target_yield = None

        row = get_crop_row(selected_crop, selected_variety)

        season = ""
        duration = 0
        rec_n = rec_p = rec_k = 0.0

        if row:
            season = row.get("Season", "") or ""
            try:
                duration = int(row.get("Duration_days", 0) or 0)
            except Exception:
                duration = 0
            rec_n = float(row.get("Rec_N", 0) or 0)
            rec_p = float(row.get("Rec_P2O5", 0) or 0)
            rec_k = float(row.get("Rec_K2O", 0) or 0)
        else:
            duration_str = request.form.get("duration", "").strip()
            try:
                duration = int(duration_str) if duration_str else 0
            except ValueError:
                duration = 0

        recommended_npk_str = f"{rec_n}:{rec_p}:{rec_k}"

        crop_session = {
            "gps_location": gps_location,
            "soil_type": soil_type,
            "crop": selected_crop,
            "variety": selected_variety,
            "season": season,
            "duration": duration,
            "recommended_npk": recommended_npk_str,
            "sowing_date": sowing_date,
            "target_yield": target_yield,
        }

        session["crop"] = crop_session
        session["stcr_rec"] = {"N": rec_n, "P2O5": rec_p, "K2O": rec_k}
        session["readings"] = []

        return redirect(url_for("readings"))

    return render_template(
        "crop_details.html",
        crop=crop_session,
        crops=CROPS,
        crop_variety_map=CROP_VARIETY_MAP,
        telangana_soils=TELANGANA_SOIL_TYPES,
    )


@app.route("/api/crop-info")
def crop_info_api():
    """AJAX endpoint: return Season, Duration, Rec NPK for selected crop+variety."""
    crop_name = request.args.get("crop", "").strip()
    variety_name = request.args.get("variety", "").strip()
    row = get_crop_row(crop_name, variety_name)
    if not row:
        return jsonify({"ok": False}), 404

    try:
        duration = int(row.get("Duration_days", 0) or 0)
    except Exception:
        duration = 0

    rec_n = float(row.get("Rec_N", 0) or 0)
    rec_p = float(row.get("Rec_P2O5", 0) or 0)
    rec_k = float(row.get("Rec_K2O", 0) or 0)

    return jsonify(
        {
            "ok": True,
            "season": row.get("Season", "") or "",
            "duration": duration,
            "rec_n": rec_n,
            "rec_p": rec_p,
            "rec_k": rec_k,
            "recommended_npk": f"{rec_n}:{rec_p}:{rec_k}",
        }
    )


@app.route("/readings", methods=["GET", "POST"])
def readings():
    readings_list = session.get("readings", [])

    if request.method == "POST":
        if "save_reading" in request.form:
            def get_float(name):
                val = request.form.get(name, "").strip()
                try:
                    return float(val) if val else 0.0
                except ValueError:
                    return 0.0

            reading = {
                "soil_n": get_float("soil_n"),
                "soil_p": get_float("soil_p"),
                "soil_k": get_float("soil_k"),
                "ec": get_float("ec"),
                "soil_temp": get_float("soil_temp"),
                "ph": get_float("ph"),
                "soil_moisture": get_float("soil_moisture"),
                "air_temp": get_float("air_temp"),
                "air_humidity": get_float("air_humidity"),
                "tds": get_float("tds"),
            }

            readings_list.append(reading)
            session["readings"] = readings_list

            if len(readings_list) >= 10:
                return redirect(url_for("calculate"))
            else:
                return redirect(url_for("readings"))

        if "reset_readings" in request.form:
            session["readings"] = []
            return redirect(url_for("readings"))

    reading_no = len(readings_list) + 1
    if reading_no > 10:
        reading_no = 10

    return render_template(
        "readings.html",
        readings=readings_list,
        reading_no=reading_no,
        max_readings=10,
    )


@app.route("/calculate")
def calculate():
    readings_list = session.get("readings", [])
    averages = compute_averages(readings_list)
    session["averages"] = averages

    crop = session.get("crop", {})
    stcr_rec = session.get("stcr_rec", {})
    soil_type = crop.get("soil_type")

    rec_n = float(stcr_rec.get("N", 0.0))
    rec_p = float(stcr_rec.get("P2O5", 0.0))
    rec_k = float(stcr_rec.get("K2O", 0.0))

    final_n = rec_n
    final_p = rec_p
    final_k = rec_k

    avg_nsoil = float(averages.get("soil_n", 0.0) or 0.0)
    avg_psoil = float(averages.get("soil_p", 0.0) or 0.0)
    avg_ksoil = float(averages.get("soil_k", 0.0) or 0.0)

    stcr_row = get_stcr_row(crop.get("crop"), crop.get("variety"))

    if stcr_row:
        target_from_farmer = crop.get("target_yield")
        if target_from_farmer is not None:
            T = float(target_from_farmer)
        else:
            try:
                T = float(stcr_row.get("TargetYield_t_ha", 0.0) or 0.0)
            except Exception:
                T = 0.0

        context = {"T": T, "SN": avg_nsoil, "SP": avg_psoil, "SK": avg_ksoil}

        stcr_n = safe_eval_formula(stcr_row.get("Formula_N", ""), context)
        stcr_p = safe_eval_formula(stcr_row.get("Formula_P2O5", ""), context)
        stcr_k = safe_eval_formula(stcr_row.get("Formula_K2O", ""), context)

        factor = soil_type_factor(soil_type)
        stcr_n *= factor
        stcr_p *= factor
        stcr_k *= factor

        stcr_n = max(0.0, stcr_n)
        stcr_p = max(0.0, stcr_p)
        stcr_k = max(0.0, stcr_k)

        final_n = min(rec_n, stcr_n) if rec_n > 0 else stcr_n
        final_p = min(rec_p, stcr_p) if rec_p > 0 else stcr_p
        final_k = min(rec_k, stcr_k) if rec_k > 0 else stcr_k

    recommended_npk_after = f"{final_n:.2f}:{final_p:.2f}:{final_k:.2f}"
    fertilizer_form = npk_to_fertilizer_form(final_n, final_p, final_k)

    calculated = {
        "recommended_npk_after": recommended_npk_after,
        "fertilizer_form": fertilizer_form,
        "final_n": final_n,
        "final_p": final_p,
        "final_k": final_k,
    }
    session["calculated"] = calculated

    irrigation_schedule = generate_irrigation_schedule(
        crop.get("sowing_date"),
        crop.get("duration", 0),
    )
    session["irrigation_schedule"] = irrigation_schedule

    fert_schedule = generate_fertilizer_schedule(
        crop.get("crop"),
        crop.get("sowing_date"),
        final_n,
        final_p,
        final_k,
    )
    session["fert_schedule"] = fert_schedule

    return redirect(url_for("report"))


@app.route("/report")
def report():
    farmer = session.get("farmer", {})
    crop = session.get("crop", {})
    readings_list = session.get("readings", [])
    averages = session.get("averages", {})
    calculated = session.get("calculated", {})
    irrigation_schedule = session.get("irrigation_schedule", [])
    fert_schedule = session.get("fert_schedule", [])

    return render_template(
        "report.html",
        farmer=farmer,
        crop=crop,
        readings=readings_list,
        averages=averages,
        calculated=calculated,
        irrigation_schedule=irrigation_schedule,
        fert_schedule=fert_schedule,
    )


# ---------- Download as PDF / Word (unchanged from earlier) ----------

@app.route("/download/pdf")
def download_pdf():
    farmer = session.get("farmer", {})
    crop = session.get("crop", {})
    averages = session.get("averages", {})
    calculated = session.get("calculated", {})
    irrigation_schedule = session.get("irrigation_schedule", [])
    fert_schedule = session.get("fert_schedule", [])

    buffer = io.BytesIO()
    p = canvas.Canvas(buffer)

    y = 800

    def line(text):
        nonlocal y
        p.drawString(50, y, text)
        y -= 15

    p.setFont("Helvetica-Bold", 14)
    line("NRSC Farmer Advisor - Soil & Crop Report")
    p.setFont("Helvetica", 11)
    y -= 10

    line(f"Farmer Name: {farmer.get('name', '')}")
    line(f"Mobile: {farmer.get('mobile', '')}")
    line(f"Village: {farmer.get('village', '')}")
    line(f"Mandal: {farmer.get('mandal', '')}")
    line(f"District: {farmer.get('district', '')}")
    y -= 10

    line(f"GPS Location: {crop.get('gps_location', '')}")
    line(f"Soil Type: {crop.get('soil_type', '')}")
    line(f"Crop: {crop.get('crop', '')}")
    line(f"Variety: {crop.get('variety', '')}")
    line(f"Season: {crop.get('season', '')}")
    line(f"Crop Duration (days): {crop.get('duration', '')}")
    line(f"Recommended N:P:K (Excel): {crop.get('recommended_npk', '')}")
    line(f"Date of Sowing: {crop.get('sowing_date', '')}")
    ty = crop.get("target_yield")
    if ty is not None:
        line(f"Target Yield for STCR (t/ha): {ty}")
    y -= 15

    line("Average Sensor Readings:")
    for label, key in [
        ("Soil Nitrogen (N)", "soil_n"),
        ("Soil Phosphorus (P)", "soil_p"),
        ("Soil Potassium (K)", "soil_k"),
        ("Electrical Conductivity (EC)", "ec"),
        ("Soil Temperature (°C)", "soil_temp"),
        ("Soil pH", "ph"),
        ("Soil Moisture (%)", "soil_moisture"),
        ("Air Temperature (°C)", "air_temp"),
        ("Air Humidity (%)", "air_humidity"),
        ("Water/Soil TDS", "tds"),
    ]:
        line(f"  {label}: {averages.get(key, 0):.2f}")
    y -= 15

    line("Final STCR-based N:P:K recommendation (kg/ha):")
    line(f"  {calculated.get('recommended_npk_after', '')}")
    y -= 10
    line("Recommended dosage in commercial fertilizer form:")
    line(f"  {calculated.get('fertilizer_form', '')}")
    y -= 15

    line("Fertilizer split schedule (basal & top dressings):")
    for fs in fert_schedule:
        line(
            f"  {fs['stage']} | {fs['das']} DAS | {fs['date']} | "
            f"N: {fs['N']} kg/ha, P2O5: {fs['P2O5']} kg/ha, K2O: {fs['K2O']} kg/ha"
        )
    y -= 15

    line("Irrigation Schedule (dates):")
    for d in irrigation_schedule:
        line(f"  {d}")

    p.showPage()
    p.save()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="NRSC_Farmer_Advisor_Report.pdf",
        mimetype="application/pdf",
    )


@app.route("/download/word")
def download_word():
    farmer = session.get("farmer", {})
    crop = session.get("crop", {})
    averages = session.get("averages", {})
    calculated = session.get("calculated", {})
    irrigation_schedule = session.get("irrigation_schedule", [])
    fert_schedule = session.get("fert_schedule", [])

    doc = Document()
    doc.add_heading("NRSC Farmer Advisor - Soil & Crop Report", level=1)

    doc.add_heading("Farmer Details", level=2)
    doc.add_paragraph(f"Name: {farmer.get('name', '')}")
    doc.add_paragraph(f"Mobile: {farmer.get('mobile', '')}")
    doc.add_paragraph(f"Village: {farmer.get('village', '')}")
    doc.add_paragraph(f"Mandal: {farmer.get('mandal', '')}")
    doc.add_paragraph(f"District: {farmer.get('district', '')}")

    doc.add_heading("Crop & Soil Details", level=2)
    doc.add_paragraph(f"GPS Location: {crop.get('gps_location', '')}")
    doc.add_paragraph(f"Soil Type: {crop.get('soil_type', '')}")
    doc.add_paragraph(f"Crop: {crop.get('crop', '')}")
    doc.add_paragraph(f"Variety: {crop.get('variety', '')}")
    doc.add_paragraph(f"Season: {crop.get('season', '')}")
    doc.add_paragraph(f"Crop Duration (days): {crop.get('duration', '')}")
    doc.add_paragraph(f"Recommended N:P:K (Excel): {crop.get('recommended_npk', '')}")
    doc.add_paragraph(f"Date of Sowing: {crop.get('sowing_date', '')}")
    ty = crop.get("target_yield")
    if ty is not None:
        doc.add_paragraph(f"Target Yield for STCR (t/ha): {ty}")

    doc.add_heading("Average Sensor Readings", level=2)
    for label, key in [
        ("Soil Nitrogen (N)", "soil_n"),
        ("Soil Phosphorus (P)", "soil_p"),
        ("Soil Potassium (K)", "soil_k"),
        ("Electrical Conductivity (EC)", "ec"),
        ("Soil Temperature (°C)", "soil_temp"),
        ("Soil pH", "ph"),
        ("Soil Moisture (%)", "soil_moisture"),
        ("Air Temperature (°C)", "air_temp"),
        ("Air Humidity (%)", "air_humidity"),
        ("Water/Soil TDS", "tds"),
    ]:
        doc.add_paragraph(f"{label}: {averages.get(key, 0):.2f}")

    doc.add_heading("Fertilizer Recommendation (STCR-based)", level=2)
    doc.add_paragraph(
        f"Final N:P:K recommendation (kg/ha): "
        f"{calculated.get('recommended_npk_after', '')}"
    )
    doc.add_paragraph(
        f"Recommended dosage in commercial fertilizer form: "
        f"{calculated.get('fertilizer_form', '')}"
    )

    doc.add_heading("Fertilizer split schedule", level=2)
    for fs in fert_schedule:
        doc.add_paragraph(
            f"{fs['stage']} | {fs['das']} DAS | {fs['date']} | "
            f"N: {fs['N']} kg/ha, P2O5: {fs['P2O5']} kg/ha, K2O: {fs['K2O']} kg/ha"
        )

    doc.add_heading("Irrigation Schedule", level=2)
    for d in irrigation_schedule:
        doc.add_paragraph(f"- {d}")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="NRSC_Farmer_Advisor_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    app.run(debug=True)
