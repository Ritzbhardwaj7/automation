import pandas as pd
import os
import re

EXCEL_PATH = r"C:\Users\admin\Desktop\TS COSTING\uploads\OB2572731 (OB2572673) @ HEATHER GREY = OB P WASHED POPOVER HENLEY.xlsx"
START_MARKER = "TARGET"

def extract_style_no(df):
    """Extract STYLE NO from any cell"""
    pattern = re.compile(r"\bO[BP]\d+\b", re.IGNORECASE)

    for _, row in df.iterrows():
        for cell in row:
            match = pattern.search(str(cell))
            if match:
                return match.group(0).upper()
    return None


def main_excel():

    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"File not found: {EXCEL_PATH}")

    raw_df = pd.read_excel(EXCEL_PATH, header=None, dtype=str)
    raw_df = raw_df.fillna("")

    # =========================
    # ✅ EXTRACT STYLE NO
    # =========================
    style_no = extract_style_no(raw_df)
    print("\n🔹 STYLE NO FOUND:", style_no)

    # FIND TARGET ROW
    start_row = None
    for i, row in raw_df.iterrows():
        if row.astype(str).str.contains(START_MARKER, case=False).any():
            start_row = i
            break

    if start_row is None:
        raise ValueError("TARGET not found in sheet.")

    # SLICE TABLE
    df = raw_df.iloc[start_row:].reset_index(drop=True)

    # REMOVE EMPTY ROWS & COLUMNS
    df = df.loc[:, (df != "").any(axis=0)]
    df = df[(df != "").any(axis=1)].reset_index(drop=True)

    # GENERIC COLUMN NAMES
    df.columns = [f"Column_{i+1}" for i in range(len(df.columns))]

    # DROP FIRST COLUMN
    if "Column_1" in df.columns:
        df = df.drop(columns=["Column_1"])

    # CLEAN EMPTY COLUMNS AGAIN
    df = df.loc[:, (df != "").any(axis=0)].reset_index(drop=True)

    # -------------------------
    # COLUMN ASSIGNMENT LOGIC
    # -------------------------
    cols = list(df.columns)

    if "Column_2" not in cols:
        raise ValueError("Description column missing")

    description_col = "Column_2"
    value_cols = [c for c in cols if c != description_col]

    # CASE 1: ONLY ONE VALUE COLUMN
    if len(value_cols) == 1:
        df = df.rename(columns={
            "Column_2": "Description",
            value_cols[0]: "ERP/BK"
        })
        df["TS"] = df["ERP/BK"]
    # CASE 2: MANY COLUMNS
    else:
        if "Column_3" in cols and "Column_4" in cols:
            df = df.rename(columns={
                "Column_2": "Description",
                "Column_3": "ERP/BK",
                "Column_4": "TS"
            })
        else:
            df = df.rename(columns={
                "Column_2": "Description",
                value_cols[-2]: "ERP/BK",
                value_cols[-1]: "TS"
            })

    # ✅ DROP ALL EXTRA COLUMNS
    df = df[["Description", "ERP/BK", "TS"]]

    # -------------------------
    # LABEL FILL DOWN + DELETE
    # -------------------------
    exact_text = "LABELS  - MAIN/ SIZE/ WASH CARE/ FTY ID/ PRICE TICKET/UPC/UCC/SENSOR LABEL.PACKING - POLYBAG / CARTON"

    rows_to_drop = []
    for i in range(len(df) - 1):
        if str(df.iloc[i]["Description"]).strip() == exact_text and str(df.iloc[i + 1]["Description"]).strip() == "":
            df.iloc[i + 1, df.columns.get_loc("Description")] = exact_text
            rows_to_drop.append(i)

    if rows_to_drop:
        df = df.drop(index=rows_to_drop).reset_index(drop=True)

    # -------------------------
    # TTL FIX
    # -------------------------
    for i in range(len(df) - 1):
        if str(df.iloc[i]["Description"]).strip() == "TTL MANUFACTURING COST":
            df.iloc[i + 1, df.columns.get_loc("Description")] = "TOTAL FOB COUNTRY"

    # -------------------------
    # REMOVE EMPTY DESCRIPTION
    # -------------------------
    df = df[df["Description"].astype(str).str.strip() != ""].reset_index(drop=True)

    # ❌ REMOVE TTL & TOTAL ROWS
    df = df[~df["Description"].str.contains(r"\bTTL\b|\bTOTAL\b", case=False, na=False)].reset_index(drop=True)

    # ✅ FORMAT TS
    def format_ts(x):
        x = str(x).strip()
        try:
            return f"{float(x):.2f}"
        except:
            return x if x else "0.00"

    df["TS"] = df["TS"].apply(format_ts)

    # ❌ DROP ERP/BK
    df = df.drop(columns=["ERP/BK"], errors="ignore")

    # ❌ REMOVE OTHER COSTING ROWS
    remove_list = ["FREIGHT COST/YD", "FABRIC PRICE/LB (FOB)", "YARN PRICE/LB"]
    df = df[~df["Description"].str.upper().isin(remove_list)].reset_index(drop=True)

    # =========================================================
    # ✅ MULTI-FABRIC COMPONENT DF (FIXED REGEX)
    # =========================================================
    components = []

    for i in range(len(df)):
        desc_raw = str(df.loc[i, "Description"])
        desc_upper = desc_raw.upper().strip()

        # Only fabric lines with a code and '@'
        # Examples:
        #  "FABRIC A - SSHF-B10 @ SHRUNKEN SHACKET FAB @ 76" (LOOP SIDE)"
        #  "FABRIC -B - SSHF-2X2-RIB B10 @ 34""
        if desc_upper.startswith("FABRIC") and "@" in desc_upper:
            # ---------- CODE ----------
            # Take text between last '-' group and first '@'
            m_code = re.search(r"-\s*([A-Z0-9\-]+(?:\s+[A-Z0-9\-]+)*)\s*@", desc_raw, flags=re.IGNORECASE)
            if m_code:
                code = m_code.group(1).strip()
            else:
                # Fallback: between last '-' and first '@'
                left, _, _ = desc_raw.partition("@")
                code = left.split("-")[-1].strip()
            code = re.sub(r"^[A-Z]\s*-\s*", "", code)


            # ---------- DESCRIPTION ----------
            parts = desc_raw.split("@")
            description = ""
            if len(parts) >= 3:
                # Between first & second '@' → base description
                base = parts[1].strip()
                tail = "@".join(parts[2:])  # width + side etc.
                # grab "(LOOP SIDE)" / "(FLAT SIDE)" if present
                m_side = re.search(r"\([^)]*\)", tail)
                side = m_side.group(0) if m_side else ""
                description = (base + " " + side).strip()
            else:
                # only one '@' → no description part, use the code
                description = code

            # Cleanup
            description = description.replace('"', "")
            description = re.sub(r"\s{2,}", " ", description)
            description = description.replace(" (", "(").strip()

            # ---------- PRICE ----------
            price = df.loc[i, "TS"]

            # ---------- YY (nearest YY row below) ----------
            yy = "0.00"
            for j in range(i + 1, len(df)):
                nxt = str(df.loc[j, "Description"]).upper().strip()
                if "YY" in nxt:
                    yy = df.loc[j, "TS"]
                    break

            components.append({
                "Code": code,
                "DESCRIPTION": description,
                "PRICE": price,
                "YY": yy
            })

    new_df = pd.DataFrame(components).reset_index(drop=True)

    # =========================================================
    # ✅ TRIM DF
    # =========================================================
    # =========================================================
# ✅ TRIM DF — Drop unwanted component labels
# =========================================================

    start_label = "CUT & MAKE"
    end_label = "OH/ WASTAGE / MARK - UP"

# Prepare Description column
    temp_desc = df["Description"].astype(str).str.upper().str.strip()

# Locate start and end indexes
    start_idx = temp_desc[temp_desc == start_label].index
    end_idx = temp_desc[temp_desc == end_label].index

# If found, extract the range
    if not start_idx.empty and not end_idx.empty:
       trim_range = df.loc[start_idx[0]:end_idx[0], ["Description", "TS"]]

    # Rename columns
       trim_df = trim_range.rename(columns={
        "Description": "components",
        "TS": "Price"
    }).reset_index(drop=True)
    else:
       trim_df = pd.DataFrame(columns=["components", "Price"])

# ---------------------------------------------------------
# ✅ DROP UNWANTED COMPONENT LABELS FROM TRIM DF
# ---------------------------------------------------------
    drop_items = [
    "CUT & MAKE",
    "WASH (TYPE)",
    "DOX / LOGISTITCS",
    "TESTING",
    "FINANCE CHARGE",
    "OH/ WASTAGE / MARK - UP"
]

    trim_df = trim_df[~trim_df["components"].str.upper().isin(drop_items)]
    trim_df = trim_df[trim_df["Price"] != "0.00"]

# Add default YY
    trim_df["YY"] = "1.000"

    return df, style_no, new_df, trim_df





# RUN
if __name__ == "__main__":

    df, style_no,new_df,trim_df = main_excel()

    if "FABRIC" in df.columns:
        df = df.drop(columns=["FABRIC"])

    print("\n===== CLEANED TABLE =====\n")
    print(df)

    print("\n===== STYLE NUMBER =====")
    print(style_no)
    print("\n===== COMPONENTS DF =====")
    print(new_df)
    print("\n===== Trim DF =====")
    print(trim_df)












































































#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@STYLE NO@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
def get_style_no():
    _, style_no, _, _ = main_excel()
    return style_no

#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#@@@@@@@@@@@@@@@@@@@@@--UPPER--@@@@@@@@@@@@@@--COSTING@--@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
def get_target_ts_value():
    df, _, _, _ = main_excel()

    # Find TARGET row
    target_row = df[df["Description"].str.contains("TARGET", case=False, na=False)]

    if target_row.empty:
        raise Exception("TARGET row not found in Excel")

    ts_value = target_row.iloc[0]["TS"]
    return ts_value

def get_notes():
    df, _, _, _ = main_excel()


    # Find TARGET row
    target_row = df[df["Description"].str.contains("FABRIC CODE #   & DESCRIPTION")]

    if target_row.empty:
        raise Exception("TARGET row not found in Excel")

    notes_value = target_row.iloc[0]["TS"]
    return notes_value
def get_labels():
    df, _, _, _ = main_excel()


    # Find TARGET row
    target_row = df[df["Description"].str.contains("LABELS  - MAIN/ SIZE/ WASH CARE/ FTY ID/ PRICE TICKET/UPC/UCC/SENSOR LABEL.PACKING - POLYBAG / CARTON")]

    if target_row.empty:
        raise Exception("TARGET row not found in Excel")

    l_value = target_row.iloc[0]["TS"]
    return l_value
def get_labour():
    df, _, _, _ = main_excel()


    # Find TARGET row
    target_row = df[df["Description"].str.contains("CUT & MAKE")]

    if target_row.empty:
        raise Exception("TARGET row not found in Excel")

    labour_value = target_row.iloc[0]["TS"]
    return labour_value
def get_wash():
    df, _, _, _ = main_excel()

    # Normalize Description column
    df["Description"] = df["Description"].astype(str).str.upper().str.strip()

    # ✅ Match only "WASH (TYPE)" row exactly
    target_row = df[df["Description"] == "WASH (TYPE)"]

    if target_row.empty:
        raise Exception("WASH (TYPE) row not found in Excel")

    wash_value = target_row.iloc[0]["TS"]
    wash_value = str(wash_value).strip()

    print("🧪 EXTRACTED WASH FROM EXCEL:", wash_value)

    return wash_value

def get_dox():
    df, _, _, _ = main_excel()


    # Normalize Description column
    df["Description"] = df["Description"].astype(str).str.upper().str.strip()

    # ✅ Match only "WASH (TYPE)" row exactly
    target_row = df[df["Description"] == "DOX / LOGISTITCS"]

    if target_row.empty:
        raise Exception("WASH (TYPE) row not found in Excel")

    dox_value = target_row.iloc[0]["TS"]
    dox_value = str(dox_value).strip()

    print("🧪 EXTRACTED WASH FROM EXCEL:", dox_value)

    return dox_value
def get_finance():
    df, _, _, _ = main_excel()

    # Normalize Description column
    df["Description"] = df["Description"].astype(str).str.upper().str.strip()

    # ✅ Match only "WASH (TYPE)" row exactly
    target_row = df[df["Description"] == "FINANCE CHARGE"]

    if target_row.empty:
        raise Exception("FINANCE CHARGE row not found in Excel")

    fin_value = target_row.iloc[0]["TS"]
    fin_value = str(fin_value).strip()

    print("🧪 EXTRACTED FINANCE CHARGE FROM EXCEL:", fin_value)

    return fin_value


def get_testing():
    df, _, _, _ = main_excel()

    # Normalize Description column
    df["Description"] = df["Description"].astype(str).str.upper().str.strip()

    # ✅ Match only "WASH (TYPE)" row exactly
    target_row = df[df["Description"] == "TESTING"]

    if target_row.empty:
        raise Exception("FINANCE CHARGE row not found in Excel")

    test_value = target_row.iloc[0]["TS"]
    test_value = str(test_value).strip()

    print("🧪 EXTRACTED testing FROM EXCEL:", test_value)

    return test_value
def get_markup():
    df, _, _, _ = main_excel()


    # Normalize Description column
    df["Description"] = df["Description"].astype(str).str.upper().str.strip()

    # ✅ Match only "WASH (TYPE)" row exactly
    target_row = df[df["Description"] == "OH/ WASTAGE / MARK - UP"]

    if target_row.empty:
        raise Exception("FINANCE CHARGE row not found in Excel")

    mark_value = target_row.iloc[0]["TS"]
    mark_value = str(mark_value).strip()

    print("🧪 EXTRACTED testing FROM EXCEL:", mark_value)

    return mark_value

#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------

