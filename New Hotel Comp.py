import streamlit as st
import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
import io

# ============================================================
# HELPERS
# ============================================================
def safe_excel_value(val):
    try:
        if pd.isna(val) or (isinstance(val, float) and (np.isnan(val) or np.isinf(val))):
            return ""
        return val
    except:
        return ""

def normalize_string(s):
    return ''.join(e for e in str(s).lower() if e.isalnum())

def fuzzy_match(val, query, threshold=90):
    if pd.isna(val):
        return False
    return fuzz.partial_ratio(str(val).lower(), str(query).lower()) >= threshold

def get_state_tax_rate(state):
    return state_tax_rates.get(state, 0)

# ============================================================
# STATE TAX RATES
# ============================================================
state_tax_rates = {
    'Alabama': 0.0039, 'Arkansas': 0.0062, 'Arizona': 0.0066, 'California': 0.0076,
    'Colorado': 0.0051, 'Connecticut': 0.0214, 'Florida': 0.0089, 'Georgia': 0.0083,
    'Iowa': 0.0157, 'Idaho': 0.0069, 'Illinois': 0.0210, 'Indiana': 0.0085,
    'Kansas': 0.0133, 'Kentucky': 0.0080, 'Louisiana': 0.0000, 'Massachusetts': 0.0112,
    'Maryland': 0.0109, 'Michigan': 0.0154, 'Missouri': 0.0097, 'Mississippi': 0.0075,
    'Montana': 0.0084, 'North Carolina': 0.0077, 'Nebraska': 0.0173, 'New Jersey': 0.0249,
    'New Mexico': 0.0080, 'Nevada': 0.0060, 'Newyork': 0.0172, 'Ohio': 0.0157,
    'Oklahoma': 0.0090, 'Oregon': 0.0097, 'Pennsylvania': 0.0158,
    'South Carolina': 0.0057, 'Tennessee': 0.0071, 'Texas': 0.0250, 'Utah': 0.0057,
    'Virginia': 0.0082, 'Washington': 0.0098
}

hotel_class_map = {
    "Budget (Low End)": 1, "Economy (Name Brand)": 2, "Midscale": 3,
    "Upper Midscale": 4, "Upscale": 5, "Upper Upscale First Class": 6,
    "Luxury Class": 7, "Independent Hotel": 8
}

# ============================================================
# STREAMLIT UI
# ============================================================
st.title("🏨 Hotel Comparable Matcher Tool")
st.write("Upload file → choose tolerance → select subjects → run → view results → download output.")

uploaded = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

# TOLERANCE MODE
reduction_mode = st.radio(
    "Market Value Tolerance Mode",
    ["Automatic (±20%)", "Manual"],
)

if reduction_mode == "Manual":
    MV_TOLERANCE = st.number_input("🔽🔼 Market Value Increase/decrease Filter %", 0.0, 500.0, 20, 1.0)
else:
    MV_TOLERANCE = 0.20

# NUMBER OF RESULTS
max_results_per_row = st.number_input(
    "Select Number of Results to Display",
    min_value=1, max_value=10, value=5
)

# ============================================================
# MAIN FILE LOAD LOGIC
# ============================================================
if uploaded is not None:

    df = pd.read_excel(uploaded)
    df.columns = [c.strip() for c in df.columns]

    for col in ["No. of Rooms", "Market Value-2024", "2024 VPR"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["No. of Rooms", "Market Value-2024", "2024 VPR"])
    df["Hotel Class Order"] = df["Hotel Class"].map(hotel_class_map)
    df = df.dropna(subset=["Hotel Class Order"])
    df["Hotel Class Order"] = df["Hotel Class Order"].astype(int)

    # ============================================================
    # SUBJECT SELECTION
    # ============================================================
    Property_Address = df['Property Address'].dropna().astype(str).str.strip().tolist()

    selected_hotels = st.multiselect(
        "🏨 Select Property Address",
        options=["[SELECT ALL]"] + Property_Address,
        default=["[SELECT ALL]"]
    )

    if "[SELECT ALL]" in selected_hotels:
        selected_rows = df.copy()
    else:
        selected_rows = df[df['Property Address'].isin(selected_hotels)]

    # RUN BUTTON
    run_process = st.button("🚀 Run Matching")

    if not run_process:
        st.warning("Click **Run Comparison** to generate results.")
        st.stop()

    # FILTER DF TO ONLY SELECTED SUBJECTS
    df_subjects = df[df["Project / Hotel Name"].isin(selected_subjects)]

    # ============================================================
    # MATCHING LOGIC
    # ============================================================
    all_results = []

    for i in range(len(df_subjects)):
        base = df_subjects.iloc[i]

        mv, vpr, rooms = base["Market Value-2024"], base["2024 VPR"], base["No. of Rooms"]
        subset = df[df.index != base.name]

        allowed = {
            1:[1,2,3],2:[1,2,3,4],3:[2,3,4,5],4:[3,4,5,6],
            5:[4,5,6,7],6:[5,6,7,8],7:[6,7,8],8:[7,8]
        }.get(base["Hotel Class Order"], [])

        mv_min, mv_max = mv * (1 - MV_TOLERANCE), mv * (1 + MV_TOLERANCE)

        mask = (
            (subset["State"] == base["State"]) &
            (subset["Property County"] == base["Property County"]) &
            (subset["No. of Rooms"] < rooms) &
            (subset["Market Value-2024"].between(mv_min, mv_max)) &
            (subset["2024 VPR"] < vpr) &
            (subset["Hotel Class Order"].isin(allowed))
        )

        matches = subset[mask]

        if matches.empty:
            all_results.append({
                "Base Property": base["Project / Hotel Name"],
                "Status": "No Match",
                "OverPaid": "",
                "Results": None
            })
            continue

        # Nearest 3
        temp = matches.copy()
        temp["dist"] = ((temp["Market Value-2024"] - mv)**2 +
                        (temp["2024 VPR"] - vpr)**2)**0.5
        nearest = temp.sort_values("dist").head(3).drop(columns="dist")

        # Remaining matches
        rem = matches.drop(nearest.index)
        least = rem.sort_values(["Market Value-2024","2024 VPR"]).head(1)
        rem = rem.drop(least.index)
        top = rem.sort_values(["Market Value-2024","2024 VPR"], ascending=[False,False]).head(1)

        final_selection = pd.concat([nearest, least, top]).head(max_results_per_row)

        # OVERPAID CALC
        median_vpr = final_selection["2024 VPR"].head(3).median()
        state_rate = get_state_tax_rate(base["State"])
        assessed = median_vpr * rooms * state_rate
        subject_tax = mv * state_rate
        overpaid = subject_tax - assessed

        all_results.append({
            "Base Property": base["Project / Hotel Name"],
            "Status": f"{len(matches)} matches found",
            "OverPaid": overpaid,
            "Results": final_selection
        })

    # ============================================================
    # DISPLAY SUMMARY
    # ============================================================
    st.subheader("📊 Final Computed Results")

    display_rows = []
    for res in all_results:
        display_rows.append({
            "Base Property": res["Base Property"],
            "Status": res["Status"],
            "OverPaid": res["OverPaid"],
        })

    summary_df = pd.DataFrame(display_rows)
    st.dataframe(summary_df)

    # DETAILS
    st.subheader("📘 Detailed Match Results")
    for res in all_results:
        st.write(f"### 🏨 {res['Base Property']}")
        st.write(f"**Status:** {res['Status']}")
        st.write(f"**OverPaid:** {res['OverPaid']}")

        if res["Results"] is not None:
            st.dataframe(res["Results"])
        st.write("---")

    # ============================================================
    # DOWNLOAD
    # ============================================================
    st.subheader("⬇️ Download Final Results")

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        for idx, res in enumerate(all_results):
            if res["Results"] is not None:
                res["Results"].to_excel(writer, sheet_name=f"Property_{idx+1}", index=False)

    st.download_button(
        label="Download Excel Results",
        data=output_buffer.getvalue(),
        file_name="comparison_results_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

