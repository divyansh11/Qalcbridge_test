import streamlit as st
import zipfile
import shutil
import xml.etree.ElementTree as ET
import pandas as pd
import re
from pathlib import Path

# -------------------------------------------
# TABLEAU → DAX MAPPING
# -------------------------------------------
TABLEAU_TO_DAX = {
    r"\bSUM\(": "SUM(",
    r"\bAVG\(": "AVERAGE(",
    r"\bMIN\(": "MIN(",
    r"\bMAX\(": "MAX(",
    r"\bCOUNT\(": "COUNT(",
    r"\bIF\s*\(": "IF(",
    r"\bTHEN\b": ",",
    r"\bELSE\b": ",",
    r"\bEND\b": ")",
    r"\bZN\(": "COALESCE(",
    r"\bISNULL\(": "ISBLANK(",
    r"\bTODAY\(\)": "TODAY()",
    r"\bNOW\(\)": "NOW()"
}

UNSUPPORTED_KEYWORDS = [
    "FIXED", "INCLUDE", "EXCLUDE", "WINDOW",
    "LOOKUP", "RUNNING", "INDEX", "RANK", "TOTAL"
]

# -------------------------------------------
# Conversion Function
# -------------------------------------------
def tableau_to_dax(tableau_formula):
    if not tableau_formula:
        return ""
    upper = tableau_formula.upper()
    if any(k in upper for k in UNSUPPORTED_KEYWORDS):
        return "⚠️ MANUAL CONVERSION REQUIRED (LOD / TABLE CALC)"
    dax = tableau_formula
    for pat, rep in TABLEAU_TO_DAX.items():
        dax = re.sub(pat, rep, dax, flags=re.IGNORECASE)
    return dax


# ---------------------------------------------------------------------
# ✅ STREAMLIT UI
# ---------------------------------------------------------------------
st.title("QalcBridge")
st.write("Upload a **.twbx** file to extract and convert all Tableau Calculated Fields to DAX.")


# ✅ Handle refresh using NEW Streamlit API
if "refresh" in st.query_params:
    st.query_params.clear()   # remove refresh flag from URL
    st.rerun()


# ✅ Maintain a dynamic key for file uploader (to reset it fully)
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

uploader_key = st.session_state["uploader_key"]

uploaded_file = st.file_uploader(
    "Upload TWBX File", 
    type=["twbx"], 
    key=f"file_uploader_{uploader_key}"
)


# ---------------------------------------------------------------------
# ✅ PROCESS THE UPLOADED FILE
# ---------------------------------------------------------------------
if uploaded_file:

    # ✅ Ensure clean extraction directory
    extract_dir = Path("extracted_twbx")
    if extract_dir.exists():
        shutil.rmtree(extract_dir)
    extract_dir.mkdir(exist_ok=True)

    st.success(f"✅ Uploaded: {uploaded_file.name}")

    # ✅ Save uploaded TWBX locally
    twbx_path = extract_dir / uploaded_file.name
    with open(twbx_path, "wb") as f:
        f.write(uploaded_file.read())

    # ✅ Extract TWBX
    with zipfile.ZipFile(twbx_path, "r") as z:
        z.extractall(extract_dir)
    st.info("✅ TWBX Extracted Successfully")

    # ✅ Locate TWB file
    twb_files = list(extract_dir.rglob("*.twb"))
    if not twb_files:
        st.error("❌ No .twb file found inside the .twbx package")
        st.stop()

    twb_path = twb_files[0]
    st.write(f"✅ Found TWB: **{twb_path.name}**")

    # ✅ Parse TWB XML
    tree = ET.parse(twb_path)
    root = tree.getroot()
    rows = []

    for datasource in root.iter("datasource"):
        ds_name = (
            datasource.attrib.get("caption")
            or datasource.attrib.get("name")
            or "Unknown Datasource"
        )

        for column in datasource.iter("column"):
            calc = column.find("calculation")
            if calc is not None:
                formula = calc.attrib.get("formula")
                rows.append({
                    "Datasource": ds_name,
                    "Calculated Field Name": (
                        column.attrib.get("caption") or column.attrib.get("name")
                    ),
                    "Tableau Formula": formula,
                    "Power BI Formula (DAX)": tableau_to_dax(formula)
                })

    df = pd.DataFrame(rows)
    st.success(f"✅ Extracted {len(df)} calculated fields")

    # ✅ Show dataframe
    st.dataframe(df, use_container_width=True)


    # ---------------------------------------------------------------------
    # ✅ CREATE EXCEL OUTPUT
    # ---------------------------------------------------------------------
    file_stem = uploaded_file.name.replace(".twbx", "")
    sheet_name = f"{file_stem}_Div_Migration_Output"
    excel_path = extract_dir / f"{sheet_name}.xlsx"

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)


    # ---------------------------------------------------------------------
    # ✅ DOWNLOAD FILE + AUTO‑RESET UI
    # ---------------------------------------------------------------------
    with open(excel_path, "rb") as f:
        if st.download_button(
            label="📥 Download Conversion Output",
            data=f,
            file_name=f"{sheet_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            # ✅ Reset file uploader (force new key)
            st.session_state["uploader_key"] = uploader_key + 1

            # ✅ Trigger auto-refresh using new API
            st.query_params.update({"refresh": "1"})
            st.rerun()


    st.markdown("---")
    st.info(f"✅ Migration Completed — Sheet Name: **{sheet_name}** 🚀")
