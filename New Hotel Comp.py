# streamlit_app.py
import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import fuzz
from io import BytesIO

# ============================================================
# CONFIG
# ============================================================
MV_TOLERANCE = 0.20  # 20% range for Market Value match

# ============================================================
# SAFE VALUE FOR EXCEL
# ============================================================
def safe_excel_value(val):
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
    'Alabama': 0.0039, 'Arkansas': 0.0062, 'Arizona': 0.0066, 'California': 0.0076,
    'Colorado': 0.0051, 'Connecticut': 0.0214, 'Florida': 0.0089, 'Georgia': 0.0083,
    'Iowa': 0.0157, 'Idaho': 0.0069, 'Illinois': 0.0210, 'Indiana': 0.0085,
    'Kansas': 0.0133, 'Kentucky': 0.0080, 'Louisiana': 0.0000, 'Massachusetts': 0.0112,
    'Maryland': 0.0109, 'Michigan': 0.0154, 'Missouri': 0.0097, 'Mississippi': 0.0075,
    'Montana': 0.0084, 'North Carolina': 0.0077, 'Nebraska': 0.0173, 'New Jersey': 0.0249,
    'New Mexico': 0.0080, 'Nevada': 0.0060, 'Newyork': 0.0172, 'Ohio': 0.0157,
    'Oklahoma': 0.0090, 'Oregon': 0.0097, 'Pennsylvania': 0.0158, 'South Carolina': 0.0057,
    'Tennessee': 0.0071, 'Texas': 0.0250, 'Utah': 0.0057, 'Virginia': 0.0082,
    'Washington': 0.0098
}

def get_state_tax_rate(state):
    return state_tax_rates.get(state, 0)

# ============================================================
# HOTEL CLASS MAPPING
# ============================================================
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

# ============================================================
# MATCHING HELPERS
# ============================================================
def get_nearest_three(df, mv, vpr):
    df = df.copy()
    df["dist"] = ((df["Market Value-2024"] - mv)**2 + (df["2024 VPR"] - vpr)**2)**0.5
    return df.sort_values("dist").head(3).drop(columns="dist")

def get_least_one(df):
    return df.sort_values(["Market Value-2024", "2024 VPR"], ascending=[True, True]).head(1)

def get_top_one(df):
    return df.sort_values(["Market Value-2024", "2024 VPR"], ascending=[False, False]).head(1)

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="Hotel Comparison Tool", layout="wide")
st.title("🏨 Hotel Market Value & VPR Comparison Tool")
st.markdown("Upload your Excel file, and the app will generate comparison results with overpaid calculation.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [col.strip() for col in df.columns]

    # Convert numeric columns
    for col in ['No. of Rooms', 'Market Value-2024', '2024 VPR']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['No. of Rooms', 'Market Value-2024', '2024 VPR'])

    df["Hotel Class Order"] = df["Hotel Class"].map(hotel_class_map)
    df = df.dropna(subset=["Hotel Class Order"])
    df["Hotel Class Order"] = df["Hotel Class Order"].astype(int)

    st.write("✅ Uploaded data preview:")
    st.dataframe(df.head())

    # ============================================================
    # PROPERTY SELECTION
    # ============================================================
    Property_Address = df['Property Address'].dropna().astype(str).str.strip().tolist()

    selected_Address = st.multiselect(
        "🏨 Select Property Address",
        options=["[SELECT ALL]"] + Property_Address,
        default=["[SELECT ALL]"]
    )

    if "[SELECT ALL]" in selected_Address:
        selected_rows = df.copy()
    else:
        selected_rows = df[df['Property Address'].isin(selected_Address)]

    # ============================================================
    # TOLERANCE MODE
    # ============================================================
    reduction_mode = st.radio(
        "📊 Market Value Increase/Decrease Filter Mode",
        ["Automated (Default 20%)", "Manual"],
        horizontal=True
    )

    if reduction_mode == "Manual":
        MV_TOLERANCE = st.number_input(
            "🔽🔼 Market Value Increase/Decrease Filter (%)",
            min_value=0.0,
            max_value=500.0,
            value=20.0,
            step=1.0
        ) / 100
    else:
        MV_TOLERANCE = 0.20

    # ============================================================
    # MAX MATCHES
    # ============================================================
    max_matches = st.number_input(
        "🔢 Max Matches Per Hotel (1–10)",
        min_value=1,
        max_value=10,
        value=5,
        step=1
    )

    max_results_per_row = max_matches

    # ============================================================
    # GENERATE BUTTON
    # ============================================================
    if st.button("🚀 Run Matching"):

        # ------------------------------------------------------------
        # 🔍 SHOW "PLEASE WAIT" MESSAGE WHILE RUNNING
        # ------------------------------------------------------------
        with st.spinner("🔍 Matching hotels, please wait..."):

            result_records = []  # for summary later

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet("Comparison Results")
                writer.sheets["Comparison Results"] = worksheet

                # Formats
                header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2'})
                border = workbook.add_format({'border': 1})
                currency0 = workbook.add_format({'num_format': '$#,##0', 'border': 1})
                currency2 = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})

                match_columns = [
                    'Property Address', 'State', 'Property County', 'Project / Hotel Name',
                    'Owner Name/ LLC Name', 'No. of Rooms', 'Market Value-2024',
                    '2024 VPR', 'Hotel Class', 'Hotel Class Number'
                ]
                all_columns = list(df.columns)
                row = 0
                status_col = len(match_columns)

                # Header row
                for c, name in enumerate(match_columns):
                    worksheet.write(row, c, name, header)
                worksheet.write(row, status_col, "Matching Results Count / Status", header)
                worksheet.write(row, status_col + 1, "OverPaid", header)

                col = status_col + 2
                for r in range(1, max_results_per_row + 1):
                    for colname in all_columns:
                        clean = "Hotel Class" if colname == "Hotel Class Order" else colname
                        worksheet.write(row, col, f"Result{r}_{clean}", header)
                        col += 1
                    worksheet.write(row, col, f"Result{r}_Hotel Class Number", header)
                    col += 1
                row += 1

                # MAIN LOOP
                for i in range(len(selected_rows)):
                    base = selected_rows.iloc[i]
                    mv = base['Market Value-2024']
                    vpr = base['2024 VPR']
                    rooms = base["No. of Rooms"]
                    # Make subset exclude current row from the full df still           
                    subset = df[df.index != base.name]  # Keep using full df for matching pool

                    allowed = {
                        1:[1,2,3],2:[1,2,3,4],3:[2,3,4,5],4:[3,4,5,6],
                        5:[4,5,6,7],6:[5,6,7,8],7:[6,7,8],8:[7,8]
                    }.get(base["Hotel Class Order"], [])

                    mv_min = mv * (1 - MV_TOLERANCE)
                    mv_max = mv * (1 + MV_TOLERANCE)

                    mask = (
                        (subset['State'] == base['State']) &
                        (subset['Property County'] == base['Property County']) &
                        (subset['No. of Rooms'] < rooms) &
                        (subset['Market Value-2024'].between(mv_min, mv_max)) &
                        (subset['2024 VPR'] < vpr) &
                        (subset['Hotel Class Order'].isin(allowed))
                    )

                    matches = subset[mask].drop_duplicates(
                        subset=['Project / Hotel Name','Property Address','Owner Name/ LLC Name']
                    )

                    # Track match status for summary
                    result_records.append("Match" if not matches.empty else "No_Match_Case")

                    # Write base row
                    for c, colname in enumerate(match_columns):
                        if colname == "Hotel Class Number":
                            val = base["Hotel Class Order"]
                            worksheet.write(row, c, safe_excel_value(val), border)
                        else:
                            val = safe_excel_value(base[colname])
                            if colname == "Market Value-2024":
                                worksheet.write(row, c, val, currency0)
                            elif colname == "2024 VPR":
                                worksheet.write(row, c, val, currency2)
                            else:
                                worksheet.write(row, c, val, border)

                    if not matches.empty:
                        nearest = get_nearest_three(matches, mv, vpr)
                        rem = matches.drop(nearest.index)
                        least = get_least_one(rem)
                        rem = rem.drop(least.index)
                        top = get_top_one(rem)
                        selected = pd.concat([nearest, least, top]).head(max_matches).reset_index(drop=True)

                        worksheet.write(row, status_col, f"Total: {len(matches)} | Selected: {len(selected)}", border)

                        median_vpr = selected["2024 VPR"].head(3).median()
                        state_rate = get_state_tax_rate(base["State"])
                        assessed = median_vpr * rooms * state_rate
                        subject_tax = mv * state_rate
                        overpaid = subject_tax - assessed
                        worksheet.write(row, status_col + 1, safe_excel_value(overpaid), currency2)

                        col = status_col + 2
                        for r in range(max_matches):
                            if r < len(selected):
                                row_df = selected.iloc[r]
                                for colname in all_columns:
                                    val = safe_excel_value(row_df[colname])
                                    if colname == "Market Value-2024":
                                        worksheet.write(row, col, val, currency0)
                                    elif colname == "2024 VPR":
                                        worksheet.write(row, col, val, currency2)
                                    elif colname == "Hotel Class Order":
                                        label = next((k for k,v in hotel_class_map.items() if v == row_df[colname]), "")
                                        worksheet.write(row, col, safe_excel_value(label), border)
                                        col += 1
                                        worksheet.write(row, col, safe_excel_value(row_df[colname]), border)
                                    else:
                                        worksheet.write(row, col, val, border)
                                    col += 1
                            else:
                                for colname in all_columns:
                                    worksheet.write(row, col, "", border)
                                    col += 1
                    else:
                        worksheet.write(row, status_col, "No_Match_Case", border)
                        worksheet.write(row, status_col + 1, "", border)

                    row += 1

                worksheet.freeze_panes(1, 0)

            processed_data = output.getvalue()

        # ✅ FIXED INDENTATION HERE
        preview_df = pd.read_excel(BytesIO(processed_data))

        # ------------------------------------------------------------
        # ✔️ AFTER PROCESSING COMPLETE MESSAGE
        # ------------------------------------------------------------
        st.success("✅ Matching Completed")

        # ------------------------------------------------------------
        # ✔️ SHOW FULL EXCEL PREVIEW
        # ------------------------------------------------------------
        st.write("📊 Full Excel Output Preview:") 
        st.dataframe(preview_df)

        # ------------------------------------------------------------
        # ✔️ SUMMARY
        # ------------------------------------------------------------
        total = len(result_records)
        matches_found = result_records.count("Match")
        no_matches = result_records.count("No_Match_Case")

        st.write("🏁 **Summary**:")
        st.write(f"- Total processed: {total}")
        st.write(f"- Matches found: {matches_found}")
        st.write(f"- No matches: {no_matches}")

        # ------------------------------------------------------------
        # DOWNLOAD BUTTON
        # ------------------------------------------------------------
        st.download_button(
            label="📥 Download Processed Excel",
            data=processed_data,
            file_name="comparison_results_streamlit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
