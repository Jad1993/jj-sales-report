import streamlit as st
import csv
import io
from collections import defaultdict
import openpyxl

st.title("📊 JJ Daily Sales Report Tool")

# ---- Simple password gate ----
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.error("Incorrect password")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ---- Core logic (same as local version) ----
def find_jj_data(file_obj):
    content = file_obj.read().decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(content)))

    jj_items = []
    jj_subtotal = None
    inside_jj = False

    for row in rows:
        if not row:
            continue
        first_cell = row[0].strip()
        first_line = first_cell.split("\n")[0].strip()

        if first_line == "JJ":
            inside_jj = True
            continue

        if inside_jj:
            if first_cell == "Sub Sub Total":
                jj_subtotal = row[5].replace(",", "").replace("(", "-").replace(")", "")
                inside_jj = False
            elif first_cell.isdigit():
                item_code = row[1]
                item_name = row[2]
                qty = row[3]
                jj_items.append((item_code, item_name, qty))
            elif first_line != "" and not first_cell.isdigit():
                inside_jj = False

    return jj_subtotal, jj_items


def build_unit_sold_lines(items):
    nb_brand_counts = defaultdict(int)
    sw_count = 0
    cp_count = 0

    for code, name, qty in items:
        qty = int(qty)
        if code.startswith("NB-"):
            parts = code.split("-")
            brand = parts[1] if len(parts) > 1 else "UNKNOWN"
            nb_brand_counts[brand] += qty
        elif code.startswith("SW-"):
            sw_count += qty
        elif code.startswith("CP-"):
            cp_count += qty

    lines = []
    for brand, qty in nb_brand_counts.items():
        lines.append(f"NB {brand} X{qty}")
    if sw_count > 0:
        lines.append(f"ANTI VIRUS X{sw_count}")
    if cp_count > 0:
        lines.append(f"AEW CP X{cp_count}")

    return lines


def write_branch_to_excel(wb, daily_cell, utd_cell, unit_sold_start_cell, unit_sold_rows,
                           daily_total, monthly_total, unit_sold_lines):
    ws = wb["Sheet1"]
    ws[daily_cell] = float(daily_total)
    ws[utd_cell] = float(monthly_total)

    start_col = unit_sold_start_cell[0]
    start_row = int(unit_sold_start_cell[1:])

    for r in range(unit_sold_rows):
        ws[f"{start_col}{start_row + r}"] = None

    for i, line in enumerate(unit_sold_lines):
        ws[f"{start_col}{start_row + i}"] = line


def render_branch_section(branch_name, daily_cell, utd_cell, unit_sold_start_cell, unit_sold_rows, excel_bytes):
    st.subheader(f"🏬 {branch_name} Branch")

    daily_key = f"daily_{branch_name}"
    monthly_key = f"monthly_{branch_name}"
    pending_key = f"pending_{branch_name}"

    daily_file = st.file_uploader(f"[{branch_name}] Upload DAILY export CSV", type="csv", key=daily_key)
    monthly_file = st.file_uploader(f"[{branch_name}] Upload MONTHLY export CSV (JJ-filtered)", type="csv", key=monthly_key)

    if st.button(f"Preview {branch_name}", key=f"preview_btn_{branch_name}"):
        if daily_file and monthly_file:
            daily_total, daily_items = find_jj_data(daily_file)
            monthly_total, _ = find_jj_data(monthly_file)
            unit_sold_lines = build_unit_sold_lines(daily_items)

            st.session_state[pending_key] = {
                "daily_total": daily_total,
                "monthly_total": monthly_total,
                "unit_sold_lines": unit_sold_lines
            }

            st.write(f"**DAILY:** RM {daily_total}")
            st.write(f"**UTD:** RM {monthly_total}")
            st.write("**Unit Sold:**")
            if unit_sold_lines:
                for line in unit_sold_lines:
                    st.write(f"- {line}")
            else:
                st.write("- (none)")
        else:
            st.warning(f"Please upload both {branch_name} files first.")

    if pending_key in st.session_state:
        data = st.session_state[pending_key]
        write_branch_to_excel(
            excel_bytes, daily_cell, utd_cell, unit_sold_start_cell, unit_sold_rows,
            data["daily_total"], data["monthly_total"], data["unit_sold_lines"]
        )

    st.divider()


st.write("Upload your current Excel report, fill in each branch below, then download the updated file at the bottom.")

excel_upload = st.file_uploader("Upload your Excel report (.xlsx)", type="xlsx", key="excel_upload")

if excel_upload:
    wb = openpyxl.load_workbook(excel_upload)

    render_branch_section("ACER", "B6", "B7", "B13", 4, wb)
    render_branch_section("HP", "G6", "G7", "G13", 4, wb)
    render_branch_section("MSI", "K6", "K7", "K13", 4, wb)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="⬇️ Download Updated Excel File",
        data=output,
        file_name="JJ_SALES_REPORT_UPDATED.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload your Excel file above to get started.")