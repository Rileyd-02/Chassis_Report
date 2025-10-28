import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill

# --- Page Config ---
st.set_page_config(page_title="Subchassis Mapper", layout="wide")

# --- Custom Dark Theme Styling ---
st.markdown("""
<style>
/* Backgrounds and general theme */
[data-testid="stAppViewContainer"] {
    background-color: #0d1117;
    color: white;
}
[data-testid="stHeader"] {
    background-color: #0d1117;
}
[data-testid="stSidebar"] {
    background-color: #161b22;
    color: white;
}

/* Titles and text */
h1, h2, h3, h4, h5, h6, p, div, label {
    color: white !important;
}

/* Upload box styling */
[data-testid="stFileUploader"] section {
    background-color: #1e252f !important;
    border: 1px solid #2e3b4e !important;
    border-radius: 10px;
}
[data-testid="stFileUploader"] section div {
    color: white !important;
}

/* Buttons */
.stButton>button {
    background-color: #30363d;
    color: white;
    border: 1px solid #6e7681;
    border-radius: 8px;
    padding: 0.5em 1.2em;
    font-weight: 500;
    transition: all 0.3s ease;
}
.stButton>button:hover {
    background-color: #484f58;
    border-color: #8b949e;
    color: white;
    transform: scale(1.02);
}

/* Expanders */
.streamlit-expanderHeader {
    background-color: #161b22 !important;
    color: white !important;
}

/* Dropdowns, selects, and inputs */
[data-baseweb="select"] > div, textarea, input {
    background-color: #161b22 !important;
    color: white !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)

# --- App Title ---
st.title("📊 Subchassis Mapper Tool")

st.markdown("""
Upload your **Planning file** and **Subchassis reference report**.  
Follow the steps below to complete the mapping process.  

---
""")

# --- Step 1: Upload Planning File ---
uploaded_planning = st.file_uploader("Upload Planning Excel File", type=["xlsx"])
planning_df = None
style_col_plan = None

if uploaded_planning:
    planning_excel = pd.ExcelFile(uploaded_planning, engine="openpyxl")
    sheet_names = planning_excel.sheet_names
    selected_sheet = st.selectbox("Select Sheet from Planning File", sheet_names)
    planning_df = planning_excel.parse(selected_sheet)
    st.success(f"✅ Loaded planning sheet: {selected_sheet}")

    style_candidates = [c for c in planning_df.columns if "style" in c.lower()]
    style_col_plan = st.selectbox(
        "Select Style Column in Planning File",
        style_candidates if style_candidates else planning_df.columns
    )

# --- Step 2: Upload Subchassis Reference File ---
uploaded_sub = st.file_uploader("Upload Subchassis Reference File", type=["xlsx"])
sub_df = None

if uploaded_sub:
    sub_excel = pd.ExcelFile(uploaded_sub, engine="openpyxl")
    sheet_names_sub = sub_excel.sheet_names
    selected_sheet_sub = st.selectbox("Select Sheet from Subchassis File", sheet_names_sub)
    sub_df = sub_excel.parse(selected_sheet_sub)
    st.success(f"✅ Loaded subchassis sheet: {selected_sheet_sub}")

    style_candidates_sub = [c for c in sub_df.columns if "style" in c.lower()]
    style_col_sub = st.selectbox(
        "Select Style Column in Subchassis File",
        style_candidates_sub if style_candidates_sub else sub_df.columns
    )
    customer_col = st.selectbox("Select Customer Column", sub_df.columns)
    dept_col = st.selectbox("Select Department Column", sub_df.columns)
    season_col = st.selectbox("Select Season Column (Optional)", ["<None>"] + list(sub_df.columns))

    with st.expander("🔎 Apply Filters (Optional)"):
        customer_filter = st.multiselect(
            "Filter by Customer",
            options=sub_df[customer_col].dropna().unique(),
            default=sub_df[customer_col].dropna().unique()
        )
        dept_filter = st.multiselect(
            "Filter by Department",
            options=sub_df[dept_col].dropna().unique(),
            default=sub_df[dept_col].dropna().unique()
        )
        season_filter = None
        if season_col != "<None>":
            season_filter = st.multiselect(
                "Filter by Season",
                options=sub_df[season_col].dropna().unique(),
                default=sub_df[season_col].dropna().unique()
            )

# --- Step 3: Mapping Logic ---
if planning_df is not None and sub_df is not None and st.button("Map Subchassis"):
    try:
        planning_df[style_col_plan] = planning_df[style_col_plan].astype(str).str.strip()
        sub_df[style_col_sub] = sub_df[style_col_sub].astype(str).str.strip()

        sub_filtered = sub_df[
            sub_df[customer_col].isin(customer_filter) &
            sub_df[dept_col].isin(dept_filter)
        ]
        if season_filter is not None:
            sub_filtered = sub_filtered[sub_filtered[season_col].isin(season_filter)]

        merge_cols = [style_col_sub, "LatestSubChassis", customer_col, dept_col]
        if season_col and season_col != "<None>":
            merge_cols.append(season_col)

        merged_df = pd.merge(
            planning_df,
            sub_filtered[merge_cols],
            left_on=style_col_plan,
            right_on=style_col_sub,
            how="left"
        )

        if style_col_sub != style_col_plan:
            merged_df.drop(columns=[style_col_sub], inplace=True)

        total_styles = len(planning_df)
        matched_styles = merged_df["LatestSubChassis"].notna().sum()
        unmatched_styles = total_styles - matched_styles

        st.subheader("📋 Summary")
        st.markdown(f"""
        - **Total Styles in Planning File:** {total_styles}  
        - **Mapped Styles:** ✅ {matched_styles}  
        - **Unmapped Styles:** ❌ {unmatched_styles}  
        """)

        st.subheader("👀 Preview of Mapped Data")
        st.dataframe(merged_df.head(20))

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Mapped Data")
            ws = writer.sheets["Mapped Data"]

            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            latest_col_idx = merged_df.columns.get_loc("LatestSubChassis") + 1
            for row_idx, value in enumerate(merged_df["LatestSubChassis"], start=2):
                if pd.isna(value):
                    ws.cell(row=row_idx, column=latest_col_idx).fill = red_fill

        st.download_button(
            label="⬇️ Download Mapped Excel File",
            data=output.getvalue(),
            file_name="Mapped_Planning_Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
