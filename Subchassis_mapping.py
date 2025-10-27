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

# --- CUSTOM CSS THEME ---
st.markdown("""
<style>
/* General app background and font */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #f0f4f8 0%, #e6ebf2 100%);
    color: #1c1c1c;
    font-family: 'Segoe UI', Roboto, sans-serif;
}

/* Sidebar styling */
[data-testid="stSidebar"] {
    background-color: #1e293b;
    color: white;
}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] p {
    color: white !important;
}

/* Titles */
h1, h2, h3 {
    color: #0f172a;
}

/* Buttons */
div.stButton > button:first-child {
    background-color: #2563eb;
    color: white;
    border-radius: 8px;
    border: none;
    padding: 0.6em 1.2em;
    font-weight: 600;
    transition: background-color 0.3s ease;
}
div.stButton > button:first-child:hover {
    background-color: #1d4ed8;
}

/* Success messages */
.stSuccess {
    background-color: #dcfce7;
    border-left: 5px solid #16a34a;
    padding: 0.8em;
    border-radius: 8px;
}

/* Expander */
.streamlit-expanderHeader {
    background-color: #f1f5f9 !important;
    border-radius: 5px;
    font-weight: 600;
}

/* DataFrame styling */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1);
}
</style>
""", unsafe_allow_html=True)


# --- APP HEADER ---
st.title("📊 Subchassis Mapper Tool")
st.caption("Easily map subchassis data between Planning and Reference Excel sheets.")
st.divider()

# --- INSTRUCTIONS SECTION ---
with st.container():
    st.markdown("""
    ### 🧭 **How It Works**
    1️⃣ Upload your **Planning file**  
    2️⃣ Select the correct **Sheet and Style column**  
    3️⃣ Upload your **Subchassis Reference Report**  
    4️⃣ Choose **Customer, Department, and optional Season columns**  
    5️⃣ Click **Map Subchassis** to generate results and download the mapped Excel  

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

    # Column selectors
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


# --- STEP 3: PROCESS MAPPING ---
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
        st.metric("Total Styles", total_styles)
        st.metric("Mapped Styles", matched_styles)
        st.metric("Unmapped Styles", unmatched_styles)

        st.subheader("👁 Preview of Mapped Data")
        st.dataframe(merged_df.head(20), use_container_width=True)

        # Save with highlights for missing subchassis
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
            label="💾 Download Mapped Excel File",
            data=output.getvalue(),
            file_name="Mapped_Planning_Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
