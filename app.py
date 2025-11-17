# app.py
# Streamlit app:
# 1. Upload Excel (all columns as string)
# 2. Ask for output folder
# 3. Keep only PROGRAMSUBTYPENAME == "ADOLOSCENT"
# 4. From those, keep rows with missing School UDISE
# 5. Optional: filter to ProgramLaunchName from a second file
# 6. Export one Excel per ProgramLaunchName
# 7. Show summary table + allow ZIP download of all files

import re
import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="UDISE Missing Extractor", layout="wide")

st.title("UDISE Missing Extractor (Adolescent Only)")

st.markdown(
    """
1. Upload the raw Excel file.  
2. (Optional) Upload another Excel file that contains the list of **ProgramLaunchName** you want to filter on.  
3. Enter the folder path where you want output files to be saved.  
4. The app will:
   - Keep only rows where **PROGRAMSUBTYPENAME = "ADOLOSCENT"**  
   - From those, find rows with missing **School UDISE**  
   - If a ProgramLaunchName list is uploaded, keep only those **ProgramLaunchName**  
   - Split and save them into separate Excel files by **ProgramLaunchName**  
   - Show a summary table and let you download all files as a **ZIP**

👉 To download the files, please go to the end of the page and click on **Download ZIP (all ProgramLaunchName files)**.
"""
)

# -----------------------------
# Helper functions
# -----------------------------
def safe_filename(name: str) -> str:
    """Turn ProgramLaunchName into a safe filename."""
    if not isinstance(name, str):
        name = str(name)
    name = name.strip()
    # Keep alphanumeric, space, dash, underscore; replace others with "_"
    name = re.sub(r"[^\w\-\s]", "_", name)
    # Replace spaces with underscore
    name = re.sub(r"\s+", "_", name)
    if not name:
        name = "UnknownProgramLaunch"
    return name[:150]  # avoid extremely long names


def load_excel_as_string(file) -> pd.DataFrame:
    """Read Excel with all columns as string dtype (where possible)."""
    try:
        df = pd.read_excel(file, dtype="string", engine="openpyxl")
    except TypeError:
        # Fallback for older pandas that don't support dtype="string"
        df = pd.read_excel(file, engine="openpyxl")
        df = df.astype("string")
    return df


def get_programlaunch_list_from_file(file) -> list[str]:
    """
    Read an Excel file and extract a list of ProgramLaunchName values.

    Priority:
    1. Column whose name (after strip+lower) == 'programlaunchname'
    2. Otherwise, first column.
    """
    df_pl = load_excel_as_string(file)
    if df_pl.empty or len(df_pl.columns) == 0:
        return []

    # Try to find a column explicitly named ProgramLaunchName (case-insensitive)
    col_candidates = [
        c for c in df_pl.columns if c.strip().lower() == "programlaunchname"
    ]
    if col_candidates:
        col = col_candidates[0]
    else:
        # Fallback: use the first column
        col = df_pl.columns[0]

    values = (
        df_pl[col]
        .dropna()
        .astype(str)
        .str.strip()
    )
    values = [v for v in values if v != ""]
    return list(dict.fromkeys(values))  # unique while preserving order


# -----------------------------
# UI: Upload & output folder
# -----------------------------
uploaded_file = st.file_uploader(
    "Upload main Excel file (with delivery data)", type=["xlsx", "xls"], accept_multiple_files=False
)

filter_file = st.file_uploader(
    "Optional: Upload Excel file that contains ProgramLaunchName list to filter on",
    type=["xlsx", "xls"],
    accept_multiple_files=False,
)

output_folder_input = st.text_input(
    "Output folder path (will be created if it doesn't exist)",
    value=str(Path.cwd() / "udise_missing_output"),
)

run_button = st.button("Run and Save Files")

# -----------------------------
# Main logic
# -----------------------------
if run_button:
    if not uploaded_file:
        st.error("Please upload the main Excel file first.")
    else:
        # Load main data
        with st.spinner("Reading main Excel file..."):
            df = load_excel_as_string(uploaded_file)

        st.success(f"Main file loaded with {len(df):,} rows and {len(df.columns)} columns.")

        # Expected columns (as per your list)
        expected_columns = [
            "COUNTRYNAME",
            "REGIONNAME",
            "STATENAME",
            "DISTRICTNAME",
            "Community/School",
            "School Type",
            "School UDISE",
            "PROGRAMTYPENAME",
            "PROGRAMSUBTYPENAME",
            "Day Of Session",
            "Group Registration Date",
            "Session Timing",
            "YM NAME",
            "TMO NAME",
            "ProgramLaunchName",
            "FUNDERNAME",
            "ProjectName",
            "ProjectType",
            "GROUPID",
            "Group Status",
            "Child School Name",
            "CHILDID",
            "Intervention Year",
            "CHILDREGNO",
            "DATE OF JOINING",
            "FNAME",
            "MNAME",
            "LNAME",
            "GENDER",
            "ISDOBKNOWN",
            "DATE OF BIRTH",
            "AGE",
            "CHILDGOSCHOOL",
            "SCHOOLTYPENAME",
            "Class Of the Child Attending school",
            "CHILDDROPEDSCHOOL",
            "CLASSCHILDDROPEDSCHOOL",
            "REASONFORDROUPOUT",
            "CHILDDISABILITY",
            "DISABILITYNAME",
            "OTHERS",
            "WASPARTOFMBPROGRAM",
            "PREVIOUSCHILDREGNO",
            "REMARKS",
            "STATUS",
            "GUARDIAN",
            "P_Poverty Line(APL/BPL)",
            "CONTACTTYPE",
            "CONTACTNUMBER",
            "P_Do you Have Document?",
            "DOCUMENTTYPE",
            "DOCUMENTNO",
            "P_FName",
            "P_Age",
            "RELATION",
            "RELIGIONNAME",
            "CASTE",
            "TRIBE",
            "Previous year grade",
            "School Academic Cycle",
            "School HM/Teacher Contact",
            "Child Level",
            "Parent Consent",
        ]

        missing_cols = [c for c in expected_columns if c not in df.columns]
        if missing_cols:
            st.warning(
                "The following expected columns are missing from the uploaded main file:\n\n"
                + ", ".join(missing_cols)
            )

        # Core required columns
        required_cols = ["PROGRAMSUBTYPENAME", "School UDISE", "ProgramLaunchName"]
        if any(col not in df.columns for col in required_cols):
            st.error(
                "Required columns 'PROGRAMSUBTYPENAME', 'School UDISE', or 'ProgramLaunchName' "
                "are missing. Please check your main file."
            )
        else:
            # Clean key columns
            df["PROGRAMSUBTYPENAME"] = df["PROGRAMSUBTYPENAME"].str.strip()
            df["ProgramLaunchName"] = df["ProgramLaunchName"].astype("string").str.strip()

            # Step 1: adolescent filter
            adol = df[df["PROGRAMSUBTYPENAME"].str.upper() == "ADOLOSCENT"].copy()
            st.write(f"Rows with PROGRAMSUBTYPENAME = 'ADOLOSCENT': **{len(adol):,}**")

            # Step 2: missing School UDISE: blank or NA
            udise_col = adol["School UDISE"]
            missing_mask = udise_col.isna() | udise_col.str.strip().eq("")
            missing_df = adol[missing_mask].copy()

            st.write(
                f"Rows with missing 'School UDISE' among adolescent rows: **{len(missing_df):,}**"
            )

            if missing_df.empty:
                st.info("No rows found with missing 'School UDISE'. Nothing to export.")
            else:
                # Step 3 (optional): filter by ProgramLaunchName list from second file
                pl_filter_list = []
                if filter_file is not None:
                    with st.spinner("Reading ProgramLaunchName filter file..."):
                        pl_filter_list = get_programlaunch_list_from_file(filter_file)

                    if not pl_filter_list:
                        st.warning(
                            "The ProgramLaunchName filter file did not produce any valid values. "
                            "Proceeding without this filter."
                        )
                    else:
                        st.write(
                            f"ProgramLaunchName values from filter file: **{len(pl_filter_list):,}**"
                        )
                        missing_df = missing_df[
                            missing_df["ProgramLaunchName"].isin(pl_filter_list)
                        ].copy()
                        st.write(
                            f"Rows with missing 'School UDISE' after applying ProgramLaunchName filter: "
                            f"**{len(missing_df):,}**"
                        )

                        if missing_df.empty:
                            st.info(
                                "After applying the ProgramLaunchName filter, "
                                "no rows remain with missing UDISE. Nothing to export."
                            )
                            st.stop()

                # Summary table: ProgramLaunchName vs count of missing UDISE
                summary_df = (
                    missing_df.groupby("ProgramLaunchName", dropna=False)
                    .size()
                    .reset_index(name="Missing_UDISE_Count")
                    .sort_values("Missing_UDISE_Count", ascending=False)
                )

                st.subheader("Summary: Missing UDISE by ProgramLaunchName")
                st.dataframe(summary_df, use_container_width=True)

                # Prepare output folder
                output_dir = Path(output_folder_input).expanduser()
                try:
                    output_dir.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    st.error(f"Could not create/access output folder: {e}")
                else:
                    groups = missing_df.groupby("ProgramLaunchName", dropna=False)
                    saved_files = []

                    # Create ZIP in memory
                    zip_buffer = io.BytesIO()

                    with st.spinner("Saving Excel files by ProgramLaunchName..."):
                        with zipfile.ZipFile(
                            zip_buffer, "w", zipfile.ZIP_DEFLATED
                        ) as zf:
                            for program_name, g in groups:
                                safe_name = safe_filename(program_name)
                                file_path = output_dir / f"{safe_name}.xlsx"

                                # Save to disk
                                g.to_excel(file_path, index=False)
                                saved_files.append(str(file_path))

                                # Also add to ZIP (in-memory)
                                sub_buffer = io.BytesIO()
                                with pd.ExcelWriter(
                                    sub_buffer, engine="openpyxl"
                                ) as writer:
                                    g.to_excel(writer, index=False)
                                zf.writestr(f"{safe_name}.xlsx", sub_buffer.getvalue())

                    st.success(
                        f"Exported {len(saved_files)} Excel file(s) to:\n\n`{output_dir}`"
                    )

                    st.write("Sample of output paths:")
                    for fp in saved_files[:10]:
                        st.write("- ", fp)
                    if len(saved_files) > 10:
                        st.write(f"... and {len(saved_files) - 10} more files.")

                    # Prepare ZIP for download
                    zip_buffer.seek(0)
                    st.subheader("Download all files as ZIP")
                    st.download_button(
                        label="Download ZIP (all ProgramLaunchName files)",
                        data=zip_buffer.getvalue(),
                        file_name="udise_missing_by_programlaunch.zip",
                        mime="application/zip",
                    )
