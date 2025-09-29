import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.comments import Comment

# ---------------- Helpers ----------------
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower()) if s is not None else ""


def extract_term_abbrev_table(tag_summary_raw: pd.DataFrame) -> pd.DataFrame:
    """Extract Term ↔ Abbreviation pairs from Tag Summary (embedded labels)."""
    df = tag_summary_raw.copy().astype(str).applymap(lambda x: x.strip() if pd.notna(x) else "")
    rows, cols = df.shape
    term_pos = abbr_pos = None

    for r in range(rows):
        for c in range(cols):
            cell = _norm(df.iat[r, c])
            if cell == "term" and term_pos is None:
                term_pos = (r, c)
            if cell == "abbreviation" and abbr_pos is None:
                abbr_pos = (r, c)

    if not term_pos or not abbr_pos:
        return pd.DataFrame(columns=["Term", "Abbreviation"])

    row_start = max(term_pos[0], abbr_pos[0]) + 1
    term_col, abbr_col = term_pos[1], abbr_pos[1]

    out_rows, blanks = [], 0
    for r in range(row_start, rows):
        term_val = df.iat[r, term_col] if term_col is not None else ""
        abbr_val = df.iat[r, abbr_col] if abbr_col is not None else ""

        if not term_val and not abbr_val:
            blanks += 1
        else:
            blanks = 0
        if blanks >= 3:
            break

        if _norm(term_val) == "term" or _norm(abbr_val) == "abbreviation":
            continue

        if term_val or abbr_val:
            out_rows.append({"Term": term_val, "Abbreviation": abbr_val})

    return pd.DataFrame(out_rows).drop_duplicates().reset_index(drop=True)


def extract_name_abbrev_table(sheet_raw: pd.DataFrame) -> pd.DataFrame:
    """Extract Name ↔ Abbreviation pairs from equipment sheets (embedded labels)."""
    df = sheet_raw.copy().astype(str).applymap(lambda x: x.strip() if pd.notna(x) else "")
    rows, cols = df.shape
    name_pos = abbr_pos = None

    for r in range(rows):
        for c in range(cols):
            cell = _norm(df.iat[r, c])
            if cell == "name" and name_pos is None:
                name_pos = (r, c)
            if "abbreviation" in cell and abbr_pos is None:  # relaxed condition
                abbr_pos = (r, c)

    if not name_pos or not abbr_pos:
        return pd.DataFrame(columns=["Name", "Abbreviation"])

    row_start = max(name_pos[0], abbr_pos[0]) + 1
    name_col, abbr_col = name_pos[1], abbr_pos[1]

    out_rows, blanks = [], 0
    for r in range(row_start, rows):
        name_val = df.iat[r, name_col] if name_col is not None else ""
        abbr_val = df.iat[r, abbr_col] if abbr_col is not None else ""

        if not name_val and not abbr_val:
            blanks += 1
        else:
            blanks = 0
        if blanks >= 3:
            break

        if _norm(name_val) == "name" or "abbreviation" in _norm(abbr_val):
            continue

        if name_val or abbr_val:
            out_rows.append({"Name": name_val, "Abbreviation": abbr_val})

    return pd.DataFrame(out_rows).drop_duplicates().reset_index(drop=True)


# ---------------- Streamlit App ----------------
st.set_page_config(page_title="Multi-Equipment Nomenclature Generator", layout="wide")
st.title("🏢 Multi-Equipment Final Nomenclature Generator")

# Upload files
planon_file = st.file_uploader("📂 Upload Planon Excel", type=["xlsx"])
sys_file = st.file_uploader("📂 Upload System Workbook (Tag Summary + Equipment Sheets)", type=["xlsx"])

if planon_file and sys_file:
    # Load Planon
    df_planon = pd.read_excel(planon_file, dtype=str)
    df_planon.columns = df_planon.columns.str.strip()

    required = ["Location code", "Building code", "Floor Code", "Room code"]
    if not all(col in df_planon.columns for col in required):
        st.error(f"Planon must contain: {', '.join(required)}")
        st.stop()

    # --- Dependent Dropdowns ---
    loc_options = sorted(df_planon["Location code"].dropna().unique())
    location_code = st.selectbox("🌍 Select Location Code", loc_options)

    building_options = sorted(
        df_planon.loc[df_planon["Location code"] == location_code, "Building code"].dropna().unique()
    )
    building = st.selectbox("🏢 Select Building Code", building_options)

    floor_options = sorted(
        df_planon[
            (df_planon["Location code"] == location_code) &
            (df_planon["Building code"] == building)
        ]["Floor Code"].dropna().unique()
    )
    floor = st.selectbox("🏬 Select Floor Code", floor_options)

    room_options = sorted(
        df_planon[
            (df_planon["Location code"] == location_code) &
            (df_planon["Building code"] == building) &
            (df_planon["Floor Code"] == floor)
        ]["Room code"].dropna().unique()
    )
    room = st.selectbox("🚪 Select Room Code", room_options)

    # Load system workbook
    book = pd.ExcelFile(sys_file)

    # Detect Tag Summary
    tag_summary_sheet = None
    for s in book.sheet_names:
        if re.search(r"tag\s*summary", s, re.I):
            tag_summary_sheet = s
            break

    if not tag_summary_sheet:
        st.error("❌ No Tag Summary sheet found")
        st.stop()

    tag_raw = pd.read_excel(book, sheet_name=tag_summary_sheet, header=None, dtype=str)
    tag_df = extract_term_abbrev_table(tag_raw)

    if tag_df.empty:
        st.error("❌ Could not extract Term-Abbreviation pairs from Tag Summary")
        st.stop()

    # Multi-select equipments
    equip_terms = st.multiselect("⚙️ Select Equipments (from Tag Summary)", sorted(tag_df["Term"].dropna().unique()))

    # Asset number per equipment
    asset_numbers = {}
    for equip in equip_terms:
        asset_numbers[equip] = st.text_input(f"🔢 Enter Asset Number for {equip}", value="")

    # Generate nomenclatures
    if st.button("🚀 Generate Final Nomenclatures"):
        if not equip_terms:
            st.error("❌ Please select at least one equipment")
            st.stop()
        if any(not num for num in asset_numbers.values()):
            st.error("❌ Please enter asset numbers for all selected equipments")
            st.stop()

        all_nomenclatures = []

        for equip_term in equip_terms:
            try:
                equip_abbrev = tag_df.loc[tag_df["Term"] == equip_term, "Abbreviation"].iloc[0]
            except:
                equip_abbrev = "N/A"

            asset_number = asset_numbers.get(equip_term, "")

            if equip_term not in book.sheet_names:
                st.warning(f"⚠️ No sheet found for equipment '{equip_term}', skipping...")
                continue

            equip_raw = pd.read_excel(book, sheet_name=equip_term, header=None, dtype=str)
            equip_df = extract_name_abbrev_table(equip_raw)

            if equip_df.empty:
                st.warning(f"⚠️ Could not extract Name/Abbreviation from sheet '{equip_term}', skipping...")
                continue

            for _, row in equip_df.iterrows():
                tag_name = str(row["Name"])
                tag_abbr = str(row["Abbreviation"])

                # --- Clean Room code: keep only numbers ---
                room_clean = "".join(filter(str.isdigit, str(room))) if room else room

                # --- Clean Tag Abbreviation: remove numbers ---
                tag_abbr_clean = re.sub(r"\d+", "", str(tag_abbr))

                # --- Build Final Nomenclature (fixed format) ---
                loc_trimmed = str(location_code).replace("LOC-", "", 1) if location_code else ""

                loc_parts = loc_trimmed.split("-")
                if len(loc_parts) >= 2:
                    loc_prefix = loc_parts[0]          # AE
                    loc_site = loc_parts[1]            # ABUS2
                    loc_formatted = f"{loc_prefix}_{loc_site}"
                else:
                    loc_formatted = loc_trimmed

                bldg_parts = str(building).split("-", 1)
                bldg_trimmed = bldg_parts[1] if len(bldg_parts) > 1 else building

                prefix = f"{loc_formatted}_{bldg_trimmed}"
                equip_token = f"{equip_abbrev}{asset_number}"
                final = f"{prefix}_{floor}_{equip_token}_{room_clean}_{tag_abbr_clean}"

                all_nomenclatures.append([
                    location_code, building, floor, room_clean,
                    equip_term, equip_abbrev, tag_name, tag_abbr_clean, final
                ])

        if all_nomenclatures:
            out_df = pd.DataFrame(
                all_nomenclatures,
                columns=[
                    "Location code", "Building code", "Floor code", "Room code",
                    "Equipment Term", "Equipment Abbreviation",
                    "Name", "Tag Abbreviation", "Final Nomenclature"
                ]
            )

            # Add explanatory note
            note_text = "⚠️ Note: Room codes are trimmed to numeric parts only; sensor names ignore numeric values."
            note_df = pd.DataFrame([[note_text] + [""] * (len(out_df.columns)-1)], columns=out_df.columns)
            out_df = pd.concat([note_df, out_df], ignore_index=True)

            st.info(note_text)
            st.success("✅ Final Nomenclatures Generated")
            st.dataframe(out_df)

            # Save Excel
            output_path = "Final_Nomenclature_Output.xlsx"
            out_df.to_excel(output_path, index=False)

            # Add Excel comment
            wb = load_workbook(output_path)
            ws = wb.active
            note = "⚠️ If there is any numeric value in the TagAbbreviation then that numeric value need not be considered in final tagging."
            ws["I1"].comment = Comment(note, "System")
            wb.save(output_path)

            # Show warning in dashboard
            st.warning(note)

            # Download
            st.download_button(
                "📥 Download Excel",
                data=open(output_path, "rb").read(),
                file_name="Final_Nomenclature_Output.xlsx"
            )
        else:
            st.error("❌ No nomenclatures generated. Please check selections and sheet contents.")
