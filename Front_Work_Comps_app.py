import streamlit as st
import pandas as pd
import numpy as np
import math
import io

# =====================================================
# 🔄 EMBEDDED PROCESSING ANIMATION (NO FILES)
# =====================================================

def show_processing_animation(text="Processing files… please wait"):
    return f"""
    <style>
    .loader-wrapper {{
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 20px;
    }}
    .loader {{
        border: 6px solid #e0e0e0;
        border-top: 6px solid #058f3c;
        border-radius: 50%;
        width: 60px;
        height: 60px;
        animation: spin 1s linear infinite;
    }}
    @keyframes spin {{
        0% {{ transform: rotate(0deg); }}
        100% {{ transform: rotate(360deg); }}
    }}
    .loader-text {{
        margin-top: 12px;
        font-family: "Segoe UI", sans-serif;
        font-size: 14px;
        color: #333;
    }}
    </style>

    <div class="loader-wrapper">
        <div class="loader"></div>
        <div class="loader-text">{text}</div>
    </div>
    """

# =====================================================
# 1. HELPER FUNCTIONS
# =====================================================

def haversine(lat1, lon1, lat2, lon2):
    try:
        lat1, lon1, lat2, lon2 = map(float, [lat1, lon1, lat2, lon2])
        lon1, lat1, lon2, lat2 = map(math.radians, [lon1, lat1, lon2, lat2])
        dlon = lon2 - lon1
        dlat = lat2 - lat1
        a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
        return 2 * math.asin(math.sqrt(a)) * 3956
    except Exception:
        return 999999

def norm_class(v):
    try:
        return int(float(v))
    except Exception:
        return np.nan

def tolerance_ok(subj_val, comp_val, pct):
    if pd.isna(subj_val) or pd.isna(comp_val) or subj_val == 0:
        return False
    return abs(comp_val - subj_val) / subj_val <= pct

def get_prefix_6(val):
    if pd.isna(val):
        return ""
    return (
        str(val).lower()
        .replace(" ", "")
        .replace(".", "")
        .replace("-", "")
        .replace(",", "")
        .replace("/", "")
    )[:6]

def unique_ok(subject, candidate, chosen, is_hotel):
    def norm(x): return str(x).strip().lower()
    for c in [subject] + chosen:
        if norm(c.get("Property Account No","")) == norm(candidate.get("Property Account No","")):
            return False
        if get_prefix_6(c.get("Owner Name/ LLC Name","")) == get_prefix_6(candidate.get("Owner Name/ LLC Name","")):
            return False
        if is_hotel:
            if get_prefix_6(c.get("Hotel Name","")) == get_prefix_6(candidate.get("Hotel Name","")):
                return False
    return True

def class_ok_hotel(s, c):
    if s == 8: return c == 8
    if c == 8: return False
    if s == 7: return c in (6,7)
    if s == 6: return c in (5,6,7)
    return abs(c - s) <= 2

def class_ok_other(s, c):
    try:
        return abs(int(s) - int(c)) <= 2
    except:
        return False

# =====================================================
# 2. CORE MATCHING
# =====================================================

def find_comps(
    srow, src, *,
    is_hotel, use_hotel_class_rule,
    max_radius_miles,
    max_gap_pct_main, max_gap_pct_value, max_gap_pct_size,
    max_comps, use_strict_distance, use_county_match, sort_mode
):

    metric = "VPR" if is_hotel else "VPU"
    size = "Rooms" if is_hotel else "GBA"
    value = "Market Value-2023" if is_hotel else "Total Market value-2023"

    subj_metric = srow.get(metric)
    subj_size = srow.get(size)
    subj_value = srow.get(value)
    subj_class = srow.get("Class_Num")

    if pd.isna(subj_metric):
        return []

    comps = []

    for _, crow in src.iterrows():
        comp_metric = crow.get(metric)
        comp_size = crow.get(size)
        comp_value = crow.get(value)

        if pd.isna(comp_metric) or comp_metric > subj_metric:
            continue

        if not tolerance_ok(subj_metric, comp_metric, max_gap_pct_main):
            continue
        if not tolerance_ok(subj_value, comp_value, max_gap_pct_value):
            continue
        if not tolerance_ok(subj_size, comp_size, max_gap_pct_size):
            continue

        dist = haversine(srow.get("lat"), srow.get("lon"), crow.get("lat"), crow.get("lon"))

        comps.append((crow, dist, subj_metric - comp_metric))

    comps.sort(key=lambda x: (x[1], -x[2]))

    final, chosen = [], []
    for crow, dist, diff in comps:
        if unique_ok(srow, crow, chosen, is_hotel):
            r = crow.copy()
            r["Distance_Calc"] = dist
            r[f"{metric}_Diff"] = diff
            final.append(r)
            chosen.append(r)
        if len(final) == max_comps:
            break

    return final

# =====================================================
# 3. STREAMLIT APP
# =====================================================

st.set_page_config(page_title="Comp Matcher", layout="wide")
st.title("🏙️ Property Tax / Hotel Comp Matcher")

st.sidebar.header("⚙️ Configuration")

prop_type = st.sidebar.radio("Property Type", ["Hotel Property", "Other Property"])
is_hotel = prop_type == "Hotel Property"
use_hotel_class_rule = st.sidebar.checkbox("Use Hotel Class Rule", value=True)

max_radius = st.sidebar.number_input("Max Radius (Miles)", 15.0)
max_gap_pct_main = st.sidebar.slider("Max Metric Gap %", 0, 100, 50) / 100
max_gap_pct_value = st.sidebar.slider("Max Value Gap %", 0, 100, 50) / 100
max_gap_pct_size = st.sidebar.slider("Max Size Gap %", 0, 100, 50) / 100
max_comps = st.sidebar.number_input("Max Comps", 1, 10, 3)

st.markdown("### Upload Files")

col1, col2 = st.columns(2)
with col1:
    subj_file = st.file_uploader("Subject Excel", type="xlsx")
with col2:
    src_file = st.file_uploader("Source Excel", type="xlsx")

# =====================================================
# 🚀 PROCESSING WITH LOADER
# =====================================================

if subj_file and src_file:
    if st.button("🚀 Run Matching", type="primary"):

        loader = st.empty()
        status = st.empty()
        progress = st.empty()

        loader.markdown(show_processing_animation(), unsafe_allow_html=True)

        try:
            status.info("📥 Reading Excel files")
            subj = pd.read_excel(subj_file)
            src = pd.read_excel(src_file)

            status.info("🧹 Cleaning data")
            subj.columns = subj.columns.str.strip()
            src.columns = src.columns.str.strip()

            for df in (subj, src):
                if "Hotel class values" in df.columns:
                    df["Class_Num"] = df["Hotel class values"].apply(norm_class)
                df["lat"] = pd.to_numeric(df.get("lat"), errors="coerce")
                df["lon"] = pd.to_numeric(df.get("lon"), errors="coerce")

            status.info("⚙️ Running matching engine")
            results = []
            prog = progress.progress(0)

            for i, (_, srow) in enumerate(subj.iterrows()):
                loader.markdown(
                    show_processing_animation(f"Matching subject {i+1} of {len(subj)}"),
                    unsafe_allow_html=True
                )

                comps = find_comps(
                    srow, src,
                    is_hotel=is_hotel,
                    use_hotel_class_rule=use_hotel_class_rule,
                    max_radius_miles=max_radius,
                    max_gap_pct_main=max_gap_pct_main,
                    max_gap_pct_value=max_gap_pct_value,
                    max_gap_pct_size=max_gap_pct_size,
                    max_comps=max_comps,
                    use_strict_distance=True,
                    use_county_match=True,
                    sort_mode="Distance Priority"
                )

                row = srow.to_dict()
                for idx, c in enumerate(comps):
                    row[f"Comp{idx+1}_Account"] = c.get("Property Account No")
                results.append(row)

                prog.progress((i + 1) / len(subj))

            loader.empty()
            progress.empty()
            status.success("✅ Matching completed")

            df = pd.DataFrame(results)
            st.dataframe(df.head())

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                df.to_excel(w, index=False)

            st.download_button(
                "📥 Download Results",
                buf.getvalue(),
                "Automated_Comps.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            loader.empty()
            status.error(str(e))
else:
    st.info("Upload both files to begin.")
