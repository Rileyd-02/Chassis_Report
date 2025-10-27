import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Subchassis Mapper",
    layout="wide",
    page_icon="📊",
)

# --- DARK THEME CSS ---
st.markdown("""
<style>
/* Background and text */
[data-testid="stAppViewContainer"] {
    background-color: #0e1117;
    color: #f5f5f5;
    font-family: 'Segoe UI', Roboto, sans-serif;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #1a1d23;
    color: #f5f5f5;
}

/* Headings */
h1, h2, h3, h4 {
    color: #ffffff !important;
    font-weight: 600;
}

/* General text */
p, label, span, div {
    color: #e5e5e5 !important;
}

/* Buttons */
div.stButton > button:first-child {
    background: linear-gradient(90deg, #2c2f36, #3b3f47);
    color: #ffffff;
    border-radius: 8px;
    border: none;
    padding: 0.6em 1.4em;
    font-weight: 600;
    font-size: 1em;
    transition: all 0.3s ease;
}
div.stButton > button:first-child:hover {
    background: linear-gradient(90deg, #3f434b, #4a4f57);
    transform: scale(1.03);
}

/* Expanders */
.streamlit-expanderHeader {
    background-color: #1b1f25 !important;
    color: #ffffff !important;
    font-weight: 500;
    border-radius: 5px;
}

/* Success and error boxes */
.stSuccess {
    background-color: rgba(56, 178, 172, 0.1);
    border-left: 4px solid #38b2ac;
    border-radius: 6px;
}
.stError {
    background-color: rgba(255, 82, 82, 0.1);
    border-left: 4px solid #ff5252;
    border-radius: 6px;
}

/* DataFrames */
[data-testid="stDataFrame"] {
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(255,255,255,0.05);
    background-color: #16191f;
}

/* Divider */
hr {
    border: 1px solid #2a2e35;
}
</style>
""", unsafe_allow_html=True)


# --- HEADER ---
st.title("📊 Subchassis Mapper Tool")
st.caption("Dark mode enabled — a clean and professional data mapping tool.")
st.divider()

# --- INSTRUCTIONS ---
st.markdown("""
### 🧭 **How to Use**
1️⃣ Upload your **Planning File**  
2️⃣ Select **Sheet** and **Style Column**  
3️⃣ Upload your **Subchassis Reference Report**  
4️⃣ Choose **Customer**, **Department**, and optionally **Season**  
5️⃣ Click **Map Subchassis** to generate and download the results  

---
""")


# --- STEP 1: UPLOAD PLANNING FILE ---
st.subheader("🗂 Step 1: Upload Planning File")
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


# --- STEP 2: UPLOAD SUBCHASSIS FILE ---
st.subheader("📘 Step 2: Upload Subchassis Reference File")
uploaded_sub = st.file_uploader("Upload Subchassis Reference File", type=["xlsx"])
sub_df = None

if uploaded_sub:
    sub_excel = pd.ExcelFile(uploaded_sub, engine="openpyxl")
    sheet_names_sub = sub_excel.sheet_names
    selected_sheet_sub = st.selectbox("Select Sheet from Subchassis File", sheet_names_sub)
    sub_df = sub_excel.parse(selected_sheet_sub)
    st.success(f"✅ Loaded subchassis sheet: {selected_sheet_sub}")

    style_candidates_sub = [c for c in sub_df.columns if "style" in c.lower()]
    style_col_sub = st.selectbox("Select Style Column", style_candidates_sub if style_candidates_sub else sub_df.columns)
    customer_col = st.selectbox("Select Customer Column", sub_df.columns)
    dept_col = st.selectbox("Select Department Column", sub_df.columns)
    season_col = st.selectbox("Select Season Column (Optional)", ["<None>"] + list(sub_df.columns))

    # --- FILTERS ---
    with st.expander("🔍 Apply Filters (Optional)"):
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


# --- STEP 3: MAP PROCESS ---
if planning_df is not None and sub_df is not None and st.button("🚀 Map Subchassis"):
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

        st.subheader("📈 Mapping Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Styles", total_styles)
        col2.metric("Mapped", matched_styles)
        col3.metric("Unmapped", unmatched_styles)

        st.subheader("👁 Preview of Mapped Data")
        st.dataframe(merged_df.head(20), use_container_width=True)

        # Save and highlight missing values
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
