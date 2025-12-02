import streamlit as st
import pandas as pd
import numpy as np
import os
from fuzzywuzzy import fuzz
import io

# ============================================================
# CONFIG — TOLERANCE SETTINGS
# ============================================================
MV_TOLERANCE = 0.20   # Default 20% range for Market Value match (–0.2 to +0.2)

# ============================================================
# SAFE VALUE FOR EXCEL (Fixes NaN/INF problem)
# ============================================================
def safe_excel_value(val):
    """Convert invalid Excel values (NaN/inf) into empty strings."""
    try:
        if pd.isna(val) or (isinstance(val, float) and (np.isnan(val) or np.isinf(val))):
            return ""
        return val
    except:
        return ""

# ============================================================
# STRING NORMALIZATION
# ============================================================
def normalize_string(s):
    return ''.join(e for e in str(s).lower() if e.isalnum())

def fuzzy_match(val, query, threshold=90):
    if pd.isna(val):
        return False
    return fuzz.partial_ratio(str(val).lower(), str(query).lower()) >= threshold

# ============================================================
# STATE TAX RATES
# ============================================================
state_tax_rates = {
    'Alabama': 0.0039, 'Arkansas': 0.0062, 'Arizona': 0.0066, 'California': 0.0076, 'Colorado': 0.0051,
    'Connecticut': 0.0214, 'Florida': 0.0089, 'Georgia': 0.0083, 'Iowa': 0.0157, 'Idaho': 0.0069,
    'Illinois': 0.0210, 'Indiana': 0.0085, 'Kansas': 0.0133, 'Kentucky': 0.0080, 'Louisiana': 0.0000,
    'Massachusetts': 0.0112, 'Maryland': 0.0109, 'Michigan': 0.0154, 'Missouri': 0.0097, 'Mississippi': 0.0075,
    'Montana': 0.0084, 'North Carolina': 0.0077, 'Nebraska': 0.0173, 'New Jersey': 0.0249, 'New Mexico': 0.0080,
    'Nevada': 0.0060, 'Newyork': 0.0172, 'Ohio': 0.0157, 'Oklahoma': 0.0090, 'Oregon': 0.0097,
    'Pennsylvania': 0.0158, 'South Carolina': 0.0057, 'Tennessee': 0.0071, 'Texas': 0.0250, 'Utah': 0.0057,
    'Virginia': 0.0082, 'Washington': 0.0098
}

def get_state_tax_rate(state):
    return state_tax_rates.get(state, 0)

# ============================================================
# INPUT FILTERS
# ============================================================
PROPERTY_FILTER = None
OWNER_FILTER = None
HOTEL_FILTER = None

# ============================================================
# EXCEL FILE UPLOAD (FROM STREAMLIT)
# ============================================================
def load_data(file):
    df = pd.read_excel(file)
    df.columns = [col.strip() for col in df.columns]

    for col in ['No. of Rooms', 'Market Value-2024', '2024 VPR']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['No. of Rooms', 'Market Value-2024', '2024 VPR'])

    hotel_class_map = {
        "Budget (Low End)": 1,
        "Economy (Name Brand)": 2,
        "Midscale": 3,
        "Upper Midscale": 4,
        "Upscale": 5,
        "Upper Upscale First Class": 6,
        "Luxury Class": 7,
        "Independent Hotel": 8
    }

    df["Hotel Class Order"] = df["Hotel Class"].map(hotel_class_map)
    df = df.dropna(subset=["Hotel Class Order"])
    df["Hotel Class Order"] = df["Hotel Class Order"].astype(int)

    return df

# ============================================================
# MAIN LOGIC: Matching Logic
# ============================================================
def run_matching(df, mv_tolerance, max_matches, selected_properties):
    match_columns = [
        'Property Address', 'State', 'Property County', 'Project / Hotel Name',
        'Owner Name/ LLC Name', 'No. of Rooms', 'Market Value-2024',
        '2024 VPR', 'Hotel Class'
    ]

    all_columns = list(df.columns)
    results = []
    match_case_count = 0
    no_match_case_count = 0

    # Loop through properties and find matches
    for i in range(len(df)):
        base = df.iloc[i]
        mv = base['Market Value-2024']
        vpr = base['2024 VPR']
        rooms = base["No. of Rooms"]

        if base['Property Address'] not in selected_properties:
            continue

        subset = df[df.index != i]

        mv_min = mv * (1 - mv_tolerance)
        mv_max = mv * (1 + mv_tolerance)

        mask = (
            (subset['State'] == base['State']) &
            (subset['Property County'] == base['Property County']) &
            (subset['No. of Rooms'] < rooms) &
            (subset['Market Value-2024'].between(mv_min, mv_max)) &
            (subset['2024 VPR'] < vpr)
        )

        matches = subset[mask].drop_duplicates(
            subset=['Project / Hotel Name', 'Property Address', 'Owner Name/ LLC Name']
        )

        if not matches.empty:
            match_case_count += 1
            results.append(matches.head(max_matches))  # Collect top matches
        else:
            no_match_case_count += 1

    return results, match_case_count, no_match_case_count

# ============================================================
# STREAMLIT APP SETUP
# ============================================================
def main():
    st.title("🏨 Hotel Comparable Matcher Tool final")

    # Step 1: Upload Excel file
    uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

    if uploaded_file:
        df = load_data(uploaded_file)

        selected_hotels = st.multiselect(
        "🏨 Select Property Address",
        options=["[SELECT ALL]"] + Property_Address,
        default=["[SELECT ALL]"]
    )

    if "[SELECT ALL]" in selected_hotels:
        selected_rows = df.copy()
    else:
        selected_rows = df[df['Property Address'].isin(selected_hotels)]

        # Step 3: Market Value Filter
        reduction_mode = st.radio(
            "🔽🔼 Market Value Increase/decrease Filter",
            options=["Automated", "Manual"]
        )

        global MV_TOLERANCE

        if reduction_mode == "Manual":
            MV_TOLERANCE = st.number_input(
                "🔽🔼 Market Value Increase/decrease Filter (%)",
                min_value=0.0,
                max_value=500.0,
                value=20.0,     # Default 20%
                step=1.0
            ) / 100
        else:
            MV_TOLERANCE = 0.20

        # Step 4: Max Matches per Hotel
        max_matches = st.slider("🔢 Max Matches Per Hotel", min_value=1, max_value=10, value=5)

        # Step 5: Run Matching
        if st.button("🚀 Run Matching"):
            if uploaded_file and selected_properties:
                results, match_case_count, no_match_case_count = run_matching(df, MV_TOLERANCE, max_matches, selected_properties)

                # Display results table
                if results:
                    result_df = pd.concat(results)
                    st.write(f"Input Rows: {len(df)}")
                    st.write(f"Output Matches: {match_case_count} | No Matches: {no_match_case_count}")
                    st.write(result_df)

                    # Provide download button for results
                    csv = result_df.to_csv(index=False)
                    st.download_button("Download Full Results", csv, file_name="matching_results.csv", mime="text/csv")
                else:
                    st.write("No matches found!")
            else:
                st.write("Please upload a file and select properties.")

if __name__ == "__main__":
    main()

