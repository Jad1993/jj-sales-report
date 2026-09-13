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


# ---- Parse the combined multi-day CSV ----
def parse_combined_csv(file_obj):
    content = file_obj.read().decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(content)))

    header = rows[0]
    date_cols = {}  # date_str -> (qty_idx, amount_idx, gp_idx)
    col = 7
    while col < len(header):
        date_str = header[col].strip()
        if date_str:
            date_cols[date_str] = (col, col + 1, col + 2)
        col += 3

    TOTAL_AMOUNT_IDX = 5

    jj_items = []          # list of (item_code, item_name, qty_by_date dict)
    jj_total_amount = None
    jj_daily_by_date = {}  # date_str -> amount string
    inside_jj = False

    def clean_num(v):
        if not v:
            return ""
        return v.replace(",", "").replace("(", "-").replace(")", "")

    for row in rows[2:]:
        if not row:
            continue
        first_cell = row[0].strip()
        first_line = first_cell.split("\n")[0].strip()

        if first_line == "JJ":
            inside_jj = True
            continue

        if inside_jj:
            if first_cell == "Sub Sub Total":
                jj_total_amount = clean_num(row[TOTAL_AMOUNT_IDX]) or "0"
                for date_str, (qcol, acol, gcol) in date_cols.items():
                    amt = row[acol] if acol < len(row) else ""
                    jj_daily_by_date[date_str] = clean_num(amt) or "0"
                inside_jj = False
            elif first_cell.isdigit():
                item_code = row[1]
                item_name = row[2]
                qty_by_date = {}
                for date_str, (qcol, acol, gcol) in date_cols.items():
                    q = row[qcol] if qcol < len(row) else ""
                    qty_by_date[date_str] = q.strip()
                jj_items.append((item_code, item_name, qty_by_date))
            elif first_line != "" and not first_cell.isdigit():
                inside_jj = False

    return jj_total_amount, jj_daily_by_date, jj_items, list(date_cols.keys())


def build_unit_sold_lines_for_date(items, selected_date):
    nb_brand_counts = defaultdict(int)
    sw_count = 0
    cp_count = 0

    for code, name, qty_by_date in items:
        q = qty_by_date.get(selected_date, "").strip()
        if not q:
            continue
        try:
            qty = int(float(q))
        except ValueError:
            continue
        if qty <= 0:
            continue

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


def render_branch_section(branch_name, daily_cell, utd_cell, unit_sold_start_cell, unit_sold_rows, wb):
    st.subheader(f"🏬 {branch_name} Branch")

    file_key = f"combined_{branch_name}"
    pending_key = f"pending_{branch_name}"

    combined_file = st.file_uploader(f"[{branch_name}] Upload combined CSV export (Day 1 to today)", type="csv", key=file_key)

    if combined_file:
        jj_total_amount, jj_daily_by_date, jj_items, all_dates = parse_combined_csv(combined_file)

        dates_with_data = [d for d in all_dates if jj_daily_by_date.get(d, "0") not in ("0", "", "-0")]
        default_date = dates_with_data[-1] if dates_with_data else all_dates[-1]
        default_index = all_dates.index(default_date)

        selected_date = st.selectbox(
            f"[{branch_name}] Select report date",
            options=all_dates,
            index=default_index,
            key=f"date_select_{branch_name}"
        )

        if st.button(f"Preview {branch_name}", key=f"preview_btn_{branch_name}"):
            daily_total = jj_daily_by_date.get(selected_date, "0")
            unit_sold_lines = build_unit_sold_lines_for_date(jj_items, selected_date)

            st.session_state[pending_key] = {
                "daily_total": daily_total,
                "monthly_total": jj_total_amount,
                "unit_sold_lines": unit_sold_lines
            }

            st.write(f"**DAILY ({selected_date}):** RM {daily_total}")
            st.write(f"**UTD:** RM {jj_total_amount}")
            st.write("**Unit Sold:**")
            if unit_sold_lines:
                for line in unit_sold_lines:
                    st.write(f"- {line}")
            else:
                st.write("- (none)")
    else:
        st.info(f"Upload {branch_name}'s combined CSV to get started.")

    if pending_key in st.session_state:
        data = st.session_state[pending_key]
        write_branch_to_excel(
            wb, daily_cell, utd_cell, unit_sold_start_cell, unit_sold_rows,
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