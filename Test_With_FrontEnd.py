import tempfile
import streamlit as st
import io
import pandas as pd
import openpyxl
from datetime import datetime
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import re
import glob
import os
from openpyxl import load_workbook


def extract_far_dates(ws):
    """
    Extract year-end date from cell A2 and period-end date from cell C1.
    """
    try:
        year_end_cell = ws['A2'].value
        period_cell = ws['C1'].value

        year_end_date = pd.to_datetime(str(year_end_cell).split("Year End -")[-1].strip(), dayfirst=True)
        period_end_date = None
        if period_cell:
            match = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)[\' ](\d{2,4})', str(period_cell))
            if match:
                period_str = match.group(0).replace("'", " 20")
                period_end_date = pd.to_datetime("01 " + period_str, dayfirst=True) + pd.offsets.MonthEnd(0)
        return year_end_date, period_end_date
    except Exception as e:
        raise ValueError(f"Failed to extract FAR dates: {e}")

# Utility Functions
def apply_tax_component_formatting(ws):
    start_row, end_row = 15, 27
    start_col, end_col = 7, 9  # G=7, H=8, I=9
    bold_font = Font(name='Gill Sans MT', size=12, bold=True)
    regular_font = Font(name='Gill Sans MT', size=12, bold=False)
    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=start_row, column=col)
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for row in range(start_row + 1, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = regular_font
            cell.alignment = right_align if col > start_col else left_align
            cell.border = thin_border
            cell.fill = PatternFill(fill_type=None)
    for r in [22, 23]:
        for c in [8, 9]:
            cell = ws.cell(row=r, column=c)
            cell.number_format = '0.00%'
            cell.font = bold_font
            cell.alignment = right_align
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 14
    ws.column_dimensions['I'].width = 14

def format_summary_table(ws, start_row=15):
    start_col = 1
    max_row = ws.max_row
    max_col = ws.max_column
    last_row = start_row
    last_col = start_col
    for r in range(start_row, max_row+1):
        row_has_data = False
        for c in range(start_col, max_col+1):
            val = ws.cell(row=r, column=c).value
            if val is not None and str(val).strip() != "":
                row_has_data = True
                if c > last_col:
                    last_col = c
        if row_has_data:
            last_row = r
    if last_row < start_row or last_col < start_col:
        return
    font = Font(name="Gill Sans MT", size=12)
    bold_font = Font(name="Gill Sans MT", size=12, bold=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    for r in range(start_row, last_row+1):
        for c in range(start_col, last_col+1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if r == start_row:
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center", vertical="bottom")
            else:
                cell.font = font
                if isinstance(val, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="bottom")
                    if val is not None:
                        if val < 0:
                            cell.number_format = "#,##0;(#,##0)"
                        elif val == 0:
                            cell.number_format = "#,##0;-#,##0;-"
                        else:
                            cell.number_format = "#,##0"
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="bottom")
            cell.border = border
    for c in range(start_col, last_col+1):
        maxlen = 0
        for r in range(start_row, last_row+1):
            v = ws.cell(row=r, column=c).value
            vstr = str(v) if v is not None else ""
            if len(vstr) > maxlen:
                maxlen = len(vstr)
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = max(10, min(maxlen+2, 40))
    if hasattr(ws, 'sheet_view'):
        ws.sheet_view.showGridLines = False

def _month_sort_key(x):
    dt = pd.to_datetime(x, errors='coerce')
    if pd.notna(dt):
        return (0, dt)
    return (1, str(x))

# Main Processing Function (full business logic from your Test.py)
def process_far_file(file_content):

    
    # Create reusable streams from same content
    excel_for_pandas = BytesIO(file_content)
    excel_for_openpyxl = BytesIO(file_content)

    df = pd.read_excel(excel_for_pandas, sheet_name='FAR', engine='openpyxl')
    wb = openpyxl.load_workbook(excel_for_openpyxl)


    # Extract year_end_date and period_end_date from FAR sheet
    df_far_head = pd.read_excel(excel_for_pandas, sheet_name='FAR', header=None, nrows=5, engine='openpyxl')
    year_end_date, period_end_date = None, None
    for i in range(5):
        row_vals = df_far_head.iloc[i].astype(str).tolist()
        for val in row_vals:
            if 'year end' in val.lower():
                match = re.search(r'(\d{1,2} [A-Za-z]+ \d{4})', val)
                if match:
                    try:
                        year_end_date = datetime.strptime(match.group(1), '%d %B %Y')
                    except Exception:
                        pass
            if 'management accounts' in val.lower():
                match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', val)
                if match:
                    try:
                        period_end_date = datetime.strptime(match.group(1), '%d/%m/%Y')
                    except Exception:
                        pass
                match2 = re.search(r'QE?\s*([A-Za-z]+)[\'’]?(\d{2})', val)
                if match2:
                    try:
                        month = match2.group(1)
                        year = int('20' + match2.group(2))
                        period_end_date = datetime.strptime(f'{month} {year}', '%b %Y')
                    except Exception:
                        pass
                match3 = re.search(r'([A-Za-z]+)\s+(\d{4})', val)
                if match3:
                    try:
                        period_end_date = datetime.strptime(f'{match3.group(1)} {match3.group(2)}', '%B %Y')
                    except Exception:
                        pass

    if not year_end_date:
        raise Exception("❌ Could not extract year-end date from FAR sheet.")
    if not period_end_date:
        raise Exception("❌ Could not extract period end date (management account date) from FAR sheet.")

    # 🗓️ Compute year boundaries
    fy_end = pd.to_datetime(year_end_date)
    fy_start = fy_end - pd.DateOffset(years=1) + pd.DateOffset(days=1)
    mgmt_acct_month = pd.to_datetime(period_end_date)

    # 📂 Setup file label and temp output folder
    output_base_name = "uploaded_file"
    output_folder = tempfile.mkdtemp()

    # 📘 Load workbook again if needed
    excel_for_openpyxl.seek(0)
    wb = openpyxl.load_workbook(excel_for_openpyxl)


    # --- MODULE 1: Split Account Transactions into separate sheets ---
    if "Account Transactions" in wb.sheetnames:
        wsSource = wb["Account Transactions"]
        clientName = wsSource["A2"].value
        excludeAccounts = [
            "Freehold Property", "Leasehold Property", "Leasehold Property Depreciation", "Plant & Machinery", "Plant & Machinery Depreciation", "Bar & Kitchen Equipment", "Bar & Kitchen Equipment Depreciation", "Furniture & Fixtures", "Furniture & Fixtures Depreciation", "Motor Vehicles", "Motor Vehicles Depreciation", "Property Improvements", "Property Improvements Depreciation", "Refurbishment", "Refurbishment Depreciation", "Goodwill", "Goodwill Amortisation", "Historical Adjustment"
        ]
        lastRow = wsSource.max_row
        headerRow = None
        for row in range(1, lastRow+1):
            if str(wsSource.cell(row=row, column=1).value).strip().upper() == "DATE":
                headerRow = row
                break
        transactionStart = None
        accountName = None
        for row in range(1, lastRow+1):
            val = wsSource.cell(row=row, column=1).value
            if val is None or str(val).strip() == "":
                if row+1 <= lastRow and wsSource.cell(row=row+1, column=1).value not in [None, ""]:
                    accountName = wsSource.cell(row=row+1, column=1).value
                    if accountName in excludeAccounts:
                        continue
                    def get_valid_sheet_name(name):
                        for ch in "/\\:?*[]":
                            name = name.replace(ch, "_")
                        return name[:31]
                    newSheetName = get_valid_sheet_name(accountName)
                    if newSheetName in wb.sheetnames:
                        wsNew = wb[newSheetName]
                    else:
                        wsNew = wb.create_sheet(title=newSheetName)
                    # PART 1: ROWS 1-4
                    wsNew["A1"] = clientName
                    wsNew["A2"] = f"Year End - {year_end_date.strftime('%d %B %Y')}"
                    wsNew["A3"] = f"Management Accounts : {period_end_date.strftime('%b\'%y').capitalize()}"
                    wsNew["A4"] = accountName
                    for r in range(1, 5):
                        cell = wsNew.cell(row=r, column=1)
                        cell.font = openpyxl.styles.Font(name="Gill Sans MT", size=12, bold=True)
                        cell.alignment = openpyxl.styles.Alignment(horizontal="left", vertical="bottom")
                    # PART 2: SUMMARY TABLE
                    wsNew["A6"] = "Date"
                    wsNew["B6"] = "Details"
                    wsNew["C6"] = "Amount £"
                    for c in range(1, 4):
                        cell = wsNew.cell(row=6, column=c)
                        cell.font = openpyxl.styles.Font(name="Gill Sans MT", size=12, bold=True)
                    wsNew["A7"].value = None
                    wsNew["B7"].value = None
                    wsNew["C7"].value = None
                    wsNew["A8"] = period_end_date.strftime('%d/%m/%Y')
                    wsNew["B8"] = accountName
                    wsNew["C8"] = None
                    wsNew["A9"].value = None
                    wsNew["B9"].value = None
                    wsNew["C9"].value = None
                    border = openpyxl.styles.Border(
                        top=openpyxl.styles.Side(style='thin'),
                        bottom=openpyxl.styles.Side(style='thin')
                    )
                    for c in range(1, 4):
                        wsNew.cell(row=6, column=c).border = border
                        wsNew.cell(row=10, column=c).border = border

    # --- MODULE 2: Process Mappings sheet and apply FormatN logic ---
    if 'Mappings' in wb.sheetnames:
        ws_mappings = wb['Mappings']
        mappings = []
        for row in ws_mappings.iter_rows(min_row=2, max_row=ws_mappings.max_row, min_col=1, max_col=2, values_only=True):
            if row[0] and row[1]:
                mappings.append((str(row[0]), str(row[1])))
        for itemName, functionName in mappings:
            target_sheet = None
            for sheet_name in wb.sheetnames:
                if sheet_name not in ['Mappings', 'Account Transactions', 'P&L']:
                    ws_account = wb[sheet_name]
                    accountName = ws_account['A4'].value
                if accountName == itemName:
                    target_sheet = ws_account
                    break
            if target_sheet is not None:
                fmt = functionName.lower()
                # --- Format1 ---
                if fmt == 'format1':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    # Get the account head from A4
                    accountName = ws.cell(row=4, column=1).value
                    foundMatch = False
                    accountRow = None
                    for rowIndex in range(1, wsTrans.max_row + 1):
                        if wsTrans.cell(row=rowIndex, column=1).value == accountName:
                            foundMatch = True
                            accountRow = rowIndex
                            break
                    if not foundMatch:
                        # Optionally, print a warning or skip
                        continue
                    # Move to the next row for opening balance calculation
                    accountRow = accountRow + 1
                    openingBalance = (wsTrans.cell(row=accountRow, column=9).value or 0) - (wsTrans.cell(row=accountRow, column=8).value or 0)
                    startRow = accountRow + 1
                    accountDict = {}
                    uniqueMonths = set()
                    for rowIndex in range(startRow, wsTrans.max_row + 1):
                        val = wsTrans.cell(row=rowIndex, column=1).value
                        if val in ("Total", "Closing Balance"):
                            break
                        # Extract account name from Column R (18), starting at 8th character
                        acc_name_val = wsTrans.cell(row=rowIndex, column=18).value
                        if acc_name_val:
                            currentAccountName = str(acc_name_val)[7:] if len(str(acc_name_val)) >= 8 else str(acc_name_val)
                        else:
                            currentAccountName = ""
                        # Transaction date in Column A
                        transactionDate = wsTrans.cell(row=rowIndex, column=1).value
                        if isinstance(transactionDate, datetime):
                            monthKey = transactionDate.strftime("%b %Y")
                        elif transactionDate:
                            try:
                                dt = pd.to_datetime(transactionDate, errors='coerce')
                                if pd.notna(dt):
                                    monthKey = dt.strftime("%b %Y")
                                else:
                                    continue
                            except Exception:
                                continue
                        else:
                            continue
                        uniqueMonths.add(monthKey)
                        if currentAccountName not in accountDict:
                            accountDict[currentAccountName] = {}
                        if monthKey not in accountDict[currentAccountName]:
                            accountDict[currentAccountName][monthKey] = 0
                        val_i = wsTrans.cell(row=rowIndex, column=9).value or 0
                        val_h = wsTrans.cell(row=rowIndex, column=8).value or 0
                        accountDict[currentAccountName][monthKey] += val_i - val_h
                    # Generate the summary table
                    summaryStartRow = 15
                    ws.cell(row=summaryStartRow, column=1).value = "Account Name"
                    ws.cell(row=summaryStartRow, column=2).value = "Opening Balance"
                    # Sort months chronologically
                    sortedMonths = sorted(uniqueMonths, key=_month_sort_key)
                    for idx, month in enumerate(sortedMonths):
                        ws.cell(row=summaryStartRow, column=3 + idx).value = month
                    ws.cell(row=summaryStartRow, column=3 + len(sortedMonths)).value = "Closing Balance"
                    # Fill in the summary data for each account
                    summaryRow = summaryStartRow + 1
                    for currentAccountName in accountDict.keys():
                        ws.cell(row=summaryRow, column=1).value = currentAccountName
                        ws.cell(row=summaryRow, column=2).value = None  # Opening balance per account not tracked in VBA
                        for idx, month in enumerate(sortedMonths):
                            ws.cell(row=summaryRow, column=3 + idx).value = accountDict[currentAccountName].get(month, 0)
                        summaryRow += 1
                    # Add the total row
                    ws.cell(row=summaryRow, column=1).value = "Total"
                    ws.cell(row=summaryRow, column=2).value = openingBalance
                    for idx, month in enumerate(sortedMonths):
                        totalSum = sum(accountDict[acc].get(month, 0) for acc in accountDict)
                        ws.cell(row=summaryRow, column=3 + idx).value = totalSum
                    # Calculate and place the final closing balance in the last column of the total row
                    totalClosingBalance = openingBalance
                    for idx in range(len(sortedMonths)):
                        val = ws.cell(row=summaryRow, column=3 + idx).value or 0
                        totalClosingBalance += val
                    ws.cell(row=summaryRow, column=3 + len(sortedMonths)).value = totalClosingBalance
                    # Set headers and total row bold
                    for col in range(1, 4 + len(sortedMonths)):
                        ws.cell(row=summaryStartRow, column=col).font = openpyxl.styles.Font(bold=True)
                        ws.cell(row=summaryRow, column=col).font = openpyxl.styles.Font(bold=True)
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format2 ---
                elif fmt == 'format2':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    accountName = ws.cell(row=4, column=1).value
                    lastRow = wsTrans.max_row
                    summaryStartRow = 15
                    # Set headers from row 5 of Account Transactions
                    ws.cell(row=summaryStartRow, column=1).value = wsTrans.cell(row=5, column=1).value
                    ws.cell(row=summaryStartRow, column=2).value = wsTrans.cell(row=5, column=2).value
                    ws.cell(row=summaryStartRow, column=3).value = wsTrans.cell(row=5, column=5).value
                    ws.cell(row=summaryStartRow, column=4).value = wsTrans.cell(row=5, column=7).value
                    ws.cell(row=summaryStartRow, column=5).value = wsTrans.cell(row=5, column=8).value
                    ws.cell(row=summaryStartRow, column=6).value = wsTrans.cell(row=5, column=9).value
                    summaryStartRow += 1
                    foundAccount = False
                    sumH = 0
                    sumI = 0
                    currentRow = 1
                    for i in range(1, lastRow + 1):
                        if wsTrans.cell(row=i, column=1).value == accountName:
                            foundAccount = True
                            currentRow = i + 1
                            while currentRow <= lastRow:
                                val = wsTrans.cell(row=currentRow, column=1).value
                                if val is not None and (str(val).startswith("Total") or str(val).startswith("Closing Balance")):
                                    break
                                # Copy relevant columns
                                ws.cell(row=summaryStartRow, column=1).value = wsTrans.cell(row=currentRow, column=1).value
                                ws.cell(row=summaryStartRow, column=2).value = wsTrans.cell(row=currentRow, column=2).value
                                ws.cell(row=summaryStartRow, column=3).value = wsTrans.cell(row=currentRow, column=5).value
                                ws.cell(row=summaryStartRow, column=4).value = wsTrans.cell(row=currentRow, column=7).value
                                ws.cell(row=summaryStartRow, column=5).value = wsTrans.cell(row=currentRow, column=8).value
                                ws.cell(row=summaryStartRow, column=6).value = wsTrans.cell(row=currentRow, column=9).value
                                # Update sums
                                sumH += wsTrans.cell(row=currentRow, column=8).value or 0
                                sumI += wsTrans.cell(row=currentRow, column=9).value or 0
                                summaryStartRow += 1
                                currentRow += 1
                            break
                    if not foundAccount:
                        # Optionally, print a warning or skip
                        continue
                    # Calculate closing balance
                    closingBalance = sumI - sumH
                    closingBalanceRow = summaryStartRow + 1
                    ws.cell(row=closingBalanceRow, column=1).value = "Closing Balance"
                    if closingBalance > 0:
                        ws.cell(row=closingBalanceRow, column=5).value = closingBalance
                    else:
                        ws.cell(row=closingBalanceRow, column=6).value = abs(closingBalance)
                    # Bold the Closing Balance row
                    for col in range(1, 7):
                        ws.cell(row=closingBalanceRow, column=col).font = openpyxl.styles.Font(bold=True)
                    # Paste the absolute value of the Closing Balance in C8
                    ws.cell(row=8, column=3).value = abs(closingBalance)
                    # Add the Total row
                    totalRow = closingBalanceRow + 1
                    ws.cell(row=totalRow, column=1).value = "Total"
                    ws.cell(row=totalRow, column=5).value = sumH + (closingBalance if closingBalance > 0 else 0)
                    ws.cell(row=totalRow, column=6).value = sumI - (closingBalance if closingBalance < 0 else 0)
                    # Bold the Total row
                    for col in range(1, 7):
                        ws.cell(row=totalRow, column=col).font = openpyxl.styles.Font(bold=True)
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format3 ---
                elif fmt == 'format3':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    # Get the period end date from cell A8
                    periodEndDate = ws.cell(row=8, column=1).value
                    # Validate date
                    try:
                        periodEndDate_dt = pd.to_datetime(periodEndDate, errors='raise')
                    except Exception:
                        # Optionally, print a warning or skip
                        continue
                    # Get the account head from A4
                    accountName = ws.cell(row=4, column=1).value
                    foundMatch = False
                    accountRow = None
                    for rowIndex in range(1, wsTrans.max_row + 1):
                        if wsTrans.cell(row=rowIndex, column=1).value == accountName:
                            foundMatch = True
                            accountRow = rowIndex
                            break
                    if not foundMatch:
                        continue
                    # Move to the next row after the account header
                    accountRow = accountRow + 1
                    openingBalance = (wsTrans.cell(row=accountRow, column=9).value or 0) - (wsTrans.cell(row=accountRow, column=8).value or 0)
                    closingBalance = openingBalance
                    # Loop through transactions to calculate the closing balance
                    for transactionRow in range(accountRow + 1, wsTrans.max_row + 1):
                        val = wsTrans.cell(row=transactionRow, column=1).value
                        if val is not None and ("Total" in str(val) or "Closing Balance" in str(val)):
                            break
                        val_i = wsTrans.cell(row=transactionRow, column=9).value or 0
                        val_h = wsTrans.cell(row=transactionRow, column=8).value or 0
                        closingBalance += val_i - val_h
                    # Start populating the summary table from row 15
                    summaryStartRow = 15
                    ws.cell(row=summaryStartRow, column=1).value = "Date"
                    ws.cell(row=summaryStartRow, column=2).value = "Particular"
                    ws.cell(row=summaryStartRow, column=3).value = "£"
                    # Set headers bold
                    for col in range(1, 4):
                        ws.cell(row=summaryStartRow, column=col).font = openpyxl.styles.Font(bold=True)
                    summaryStartRow += 1
                    ws.cell(row=summaryStartRow, column=1).value = periodEndDate
                    ws.cell(row=summaryStartRow, column=2).value = "Balance as per statement"
                    ws.cell(row=summaryStartRow, column=3).value = ""
                    summaryStartRow += 1
                    ws.cell(row=summaryStartRow, column=1).value = periodEndDate
                    ws.cell(row=summaryStartRow, column=2).value = "Balance as per Xero"
                    ws.cell(row=summaryStartRow, column=3).value = closingBalance
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format4 ---

                elif fmt == 'format4':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    accountName = ws.cell(row=4, column=1).value
                    summaryRow = 15
                    lastRow = wsTrans.max_row
                    currentRow = 1
                    monthDict = {}
                    found = False
                    while currentRow <= lastRow:
                        if wsTrans.cell(row=currentRow, column=1).value == accountName:
                            openingBalance = wsTrans.cell(row=currentRow + 1, column=9).value or 0
                            currentRow = currentRow + 2
                            # Process transactions until 'Total PAYE' or 'Closing Balance'
                            while currentRow <= lastRow and wsTrans.cell(row=currentRow, column=1).value not in (None, "", "Total PAYE", "Closing Balance"):
                                def safe_float(val):
                                    try:
                                        return float(val)
                                    except (TypeError, ValueError):
                                        return 0
                                amountI = safe_float(wsTrans.cell(row=currentRow, column=9).value)
                                amountH = safe_float(wsTrans.cell(row=currentRow, column=8).value)
                                transactionDate = wsTrans.cell(row=currentRow, column=1).value
                                # Extract monthKey
                                if isinstance(transactionDate, datetime):
                                    monthKey = transactionDate.strftime("%B %Y")
                                elif transactionDate:
                                    try:
                                        dt = pd.to_datetime(transactionDate, errors='coerce')
                                        if pd.notna(dt):
                                            monthKey = dt.strftime("%B %Y")
                                        else:
                                            monthKey = str(transactionDate)[:7]
                                    except Exception:
                                        monthKey = str(transactionDate)[:7]
                                else:
                                    monthKey = ""
                                if monthKey not in monthDict:
                                    monthDict[monthKey] = {"liability": 0, "payment": 0}
                                # Liabilities: sum I if not HMRC/NEST
                                colC = wsTrans.cell(row=currentRow, column=3).value or ""
                                colE = wsTrans.cell(row=currentRow, column=5).value or ""
                                is_hmrc_nest = any(x in str(colC).upper() or x in str(colE).upper() for x in ["HMRC", "NEST"])
                                if not is_hmrc_nest:
                                    monthDict[monthKey]["liability"] += amountI
                                # Payment: sum H if HMRC/NEST
                                if is_hmrc_nest:
                                    monthDict[monthKey]["payment"] += amountH
                                # Subtract H from liabilities if Manual Journal in B
                                colB = wsTrans.cell(row=currentRow, column=2).value or ""
                                if str(colB).upper() == "MANUAL JOURNAL":
                                    monthDict[monthKey]["liability"] -= amountH
                                currentRow += 1
                            # Output summary table headers
                            ws.cell(row=summaryRow, column=1).value = "Month"
                            ws.cell(row=summaryRow, column=2).value = "Liability"
                            ws.cell(row=summaryRow, column=3).value = "Payment"
                            ws.cell(row=summaryRow, column=4).value = "Outstanding"
                            for col in range(1, 5):
                                ws.cell(row=summaryRow, column=col).font = openpyxl.styles.Font(bold=True)
                            summaryRow += 1
                            # Opening balance row
                            ws.cell(row=summaryRow, column=1).value = "Opening Balance"
                            ws.cell(row=summaryRow, column=2).value = openingBalance
                            ws.cell(row=summaryRow, column=4).value = openingBalance
                            summaryRow += 1
                            # Output each month's calculated values
                            for monthKey in monthDict:
                                liabilities = monthDict[monthKey]["liability"]
                                payment = monthDict[monthKey]["payment"]
                                outstanding = liabilities - payment
                                ws.cell(row=summaryRow, column=1).value = monthKey
                                ws.cell(row=summaryRow, column=2).value = liabilities
                                ws.cell(row=summaryRow, column=3).value = payment
                                ws.cell(row=summaryRow, column=4).value = outstanding
                                summaryRow += 1
                            # Calculate the total outstanding from the monthly entries including opening balance
                            totalOutstanding = openingBalance
                            for monthKey in monthDict:
                                totalOutstanding += monthDict[monthKey]["liability"] - monthDict[monthKey]["payment"]
                            ws.cell(row=summaryRow, column=1).value = "Outstanding Total"
                            ws.cell(row=summaryRow, column=4).value = totalOutstanding
                            found = True
                            break
                        currentRow += 1
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format5 ---
                elif fmt == 'format5':
                    # CorporationTaxSummaryTable logic
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    accountName = ws.cell(row=4, column=1).value
                    summaryRow = 15
                    lastRow = wsTrans.max_row
                    currentRow = 1
                    monthDict = {}
                    totalLiability = 0
                    totalPayment = 0
                    totalOutstanding = 0
                    found = False
                    while currentRow <= lastRow:
                        if wsTrans.cell(row=currentRow, column=1).value == accountName:
                            openingBalance = wsTrans.cell(row=currentRow + 1, column=9).value or 0
                            currentRow = currentRow + 2
                            # Process transactions
                            while currentRow <= lastRow and wsTrans.cell(row=currentRow, column=1).value not in (None, "", "Closing Balance") and "Total" not in str(wsTrans.cell(row=currentRow, column=1).value):
                                amountI = wsTrans.cell(row=currentRow, column=9).value or 0
                                amountH = wsTrans.cell(row=currentRow, column=8).value or 0
                                transactionDate = wsTrans.cell(row=currentRow, column=1).value
                                # Extract monthKey
                                if isinstance(transactionDate, datetime):
                                    monthKey = transactionDate.strftime("%B %Y")
                                elif transactionDate:
                                    try:
                                        dt = pd.to_datetime(transactionDate, errors='coerce')
                                        if pd.notna(dt):
                                            monthKey = dt.strftime("%B %Y")
                                        else:
                                            monthKey = str(transactionDate)[:7]
                                    except Exception:
                                        monthKey = str(transactionDate)[:7]
                                else:
                                    monthKey = ""
                                if monthKey not in monthDict:
                                    monthDict[monthKey] = {"totalI": 0, "totalH": 0, "payment": 0}
                                # Update monthly totals for liabilities and payments based on Column C
                                colC = wsTrans.cell(row=currentRow, column=3).value or ""
                                colB = wsTrans.cell(row=currentRow, column=2).value or ""
                                if str(colC).upper() != "HMRC":
                                    monthDict[monthKey]["totalI"] += amountI
                                if str(colB).upper() == "MANUAL JOURNAL":
                                    monthDict[monthKey]["totalI"] -= amountH
                                monthDict[monthKey]["totalH"] += amountH
                                if str(colC).upper() == "HMRC":
                                    monthDict[monthKey]["payment"] += amountH - amountI
                                currentRow += 1
                            # Output summary table headers
                            ws.cell(row=summaryRow, column=1).value = "Month"
                            ws.cell(row=summaryRow, column=2).value = "Liability"
                            ws.cell(row=summaryRow, column=3).value = "Payment"
                            ws.cell(row=summaryRow, column=4).value = "Outstanding"
                            ws.cell(row=summaryRow, column=5).value = "Payment Date"
                            for col in range(1, 6):
                                ws.cell(row=summaryRow, column=col).font = openpyxl.styles.Font(bold=True)
                            summaryRow += 1
                            # Opening balance row
                            ws.cell(row=summaryRow, column=1).value = "Opening Balance"
                            ws.cell(row=summaryRow, column=2).value = openingBalance
                            summaryRow += 1
                            # Output each month's calculated values
                            for monthKey in monthDict:
                                liabilities = monthDict[monthKey]["totalI"]
                                payment = monthDict[monthKey]["payment"]
                                outstanding = liabilities - payment
                                if outstanding < 0:
                                    outstanding = 0
                                ws.cell(row=summaryRow, column=1).value = monthKey
                                ws.cell(row=summaryRow, column=2).value = liabilities
                                ws.cell(row=summaryRow, column=3).value = payment
                                ws.cell(row=summaryRow, column=4).value = outstanding
                                totalLiability += liabilities
                                totalPayment += payment
                                summaryRow += 1
                            # Add opening balance to total liabilities
                            totalLiability += openingBalance
                            # Write the Balance row
                            totalOutstanding = totalLiability - totalPayment
                            ws.cell(row=summaryRow, column=1).value = "Balance"
                            ws.cell(row=summaryRow, column=2).value = totalLiability
                            ws.cell(row=summaryRow, column=3).value = totalPayment
                            ws.cell(row=summaryRow, column=4).value = totalOutstanding
                            ws.cell(row=8, column=3).value = totalOutstanding
                            # Number format for C8
                            ws.cell(row=8, column=3).number_format = "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"
                            found = True
                            break
                        currentRow += 1
                    # CalculateTaxComponents logic
                    wsPL = wb["P&L"] if "P&L" in wb.sheetnames else None
                    if wsPL:
                        lastRowPL = wsPL.max_row
                        netProfitBeforeTax = 0
                        depreciation = 0
                        netProfitBeforeTaxYTD = 0
                        depreciationYTD = 0
                        for row in range(1, lastRowPL + 1):
                            val = wsPL.cell(row=row, column=1).value
                            # Safely convert cell values to float, skip if not possible
                            def safe_float(cellval):
                                try:
                                    return float(cellval)
                                except (TypeError, ValueError):
                                    return 0
                            if val and "Profit after Taxation" in str(val):
                                netProfitBeforeTax += safe_float(wsPL.cell(row=row, column=2).value)
                                netProfitBeforeTaxYTD += safe_float(wsPL.cell(row=row, column=3).value)
                            if val and "Corporation Tax Expense" in str(val):
                                netProfitBeforeTax += safe_float(wsPL.cell(row=row, column=2).value)
                                netProfitBeforeTaxYTD += safe_float(wsPL.cell(row=row, column=3).value)
                            if val and "Depreciation" in str(val):
                                depreciation += safe_float(wsPL.cell(row=row, column=2).value)
                                depreciationYTD += safe_float(wsPL.cell(row=row, column=3).value)
                        netProfit = netProfitBeforeTax + depreciation
                        netProfitYTD = netProfitBeforeTaxYTD + depreciationYTD
                        # Tax rates
                        taxRateUpTo50K = 0.19
                        taxRateAbove50K = 0.265
                        # CT charge for Monthly Net Profit
                        if netProfit < 0:
                            ctChargeMonthly = 0
                        elif netProfit < 50000:
                            ctChargeMonthly = netProfit * taxRateUpTo50K
                        else:
                            ctChargeMonthly = netProfit * taxRateAbove50K
                        # CT charge for YTD Net Profit
                        if netProfitYTD < 0:
                            ctChargeYTD = 0
                        elif netProfitYTD < 50000:
                            ctChargeYTD = netProfitYTD * taxRateUpTo50K
                        else:
                            ctChargeYTD = netProfitYTD * taxRateAbove50K
                        # Get Month'YY from cell A8
                        monthYear = ws.cell(row=8, column=1).value
                        try:
                            monthYear_fmt = pd.to_datetime(monthYear).strftime("%b'%y")
                        except Exception:
                            monthYear_fmt = str(monthYear)
                        # Place headers starting from row 15
                        ws.cell(row=15, column=7).value = ""
                        ws.cell(row=15, column=8).value = monthYear_fmt
                        ws.cell(row=15, column=9).value = "YTD"
                        ws.cell(row=16, column=7).value = "Net profit before tax"
                        ws.cell(row=16, column=8).value = netProfitBeforeTax
                        ws.cell(row=16, column=9).value = netProfitBeforeTaxYTD
                        ws.cell(row=17, column=7).value = ""
                        ws.cell(row=17, column=8).value = ""
                        ws.cell(row=17, column=9).value = ""
                        ws.cell(row=18, column=7).value = "Depreciation"
                        ws.cell(row=18, column=8).value = depreciation
                        ws.cell(row=18, column=9).value = depreciationYTD
                        ws.cell(row=19, column=7).value = ""
                        ws.cell(row=19, column=8).value = ""
                        ws.cell(row=19, column=9).value = ""
                        ws.cell(row=20, column=7).value = "Net profit"
                        ws.cell(row=20, column=8).value = netProfit
                        ws.cell(row=20, column=9).value = netProfitYTD
                        ws.cell(row=21, column=7).value = ""
                        ws.cell(row=21, column=8).value = ""
                        ws.cell(row=21, column=9).value = ""
                        ws.cell(row=22, column=7).value = "Tax rate up to 50K profit"
                        ws.cell(row=22, column=8).value = taxRateUpTo50K
                        ws.cell(row=22, column=9).value = taxRateUpTo50K
                        ws.cell(row=23, column=7).value = "Tax rate above 50K profit"
                        ws.cell(row=23, column=8).value = taxRateAbove50K
                        ws.cell(row=23, column=9).value = taxRateAbove50K
                        ws.cell(row=24, column=7).value = ""
                        ws.cell(row=24, column=8).value = ""
                        ws.cell(row=24, column=9).value = ""
                        ws.cell(row=25, column=7).value = "CT charge"
                        ws.cell(row=25, column=8).value = ctChargeMonthly
                        ws.cell(row=25, column=9).value = ctChargeYTD
                        ws.cell(row=26, column=7).value = ""
                        ws.cell(row=26, column=8).value = ""
                        ws.cell(row=26, column=9).value = ""
                        ws.cell(row=27, column=7).value = "Total CT"
                        ws.cell(row=27, column=8).value = ctChargeMonthly
                        ws.cell(row=27, column=9).value = ctChargeYTD
                        # Bold row 15
                        for col in range(7, 10):
                            ws.cell(row=15, column=col).font = openpyxl.styles.Font(bold=True)
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format6 ---
                elif fmt == 'format6':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    accountName = ws.cell(row=4, column=1).value
                    summaryRow = 15
                    lastRow = wsTrans.max_row
                    currentRow = 1
                    monthDict = {}
                    totalLiability = 0
                    totalPayment = 0
                    totalOutstanding = 0
                    found = False
                    while currentRow <= lastRow:
                        if wsTrans.cell(row=currentRow, column=1).value == accountName:
                            openingBalance = wsTrans.cell(row=currentRow + 1, column=9).value or 0
                            currentRow = currentRow + 2
                            # Process transactions
                            while currentRow <= lastRow and wsTrans.cell(row=currentRow, column=1).value not in (None, "", "Closing Balance") and "Total" not in str(wsTrans.cell(row=currentRow, column=1).value):
                                amountI = wsTrans.cell(row=currentRow, column=9).value or 0
                                amountH = wsTrans.cell(row=currentRow, column=8).value or 0
                                transactionDate = wsTrans.cell(row=currentRow, column=1).value
                                # Extract monthKey
                                if isinstance(transactionDate, datetime):
                                    monthKey = transactionDate.strftime("%B %Y")
                                elif transactionDate:
                                    try:
                                        dt = pd.to_datetime(transactionDate, errors='coerce')
                                        if pd.notna(dt):
                                            monthKey = dt.strftime("%B %Y")
                                        else:
                                            monthKey = str(transactionDate)[:7]
                                    except Exception:
                                        monthKey = str(transactionDate)[:7]
                                else:
                                    monthKey = ""
                                if monthKey not in monthDict:
                                    monthDict[monthKey] = {"totalI": 0, "totalH": 0}
                                # Update monthly totals for liabilities based on Column C (excluding "HMRC")
                                colC = wsTrans.cell(row=currentRow, column=3).value or ""
                                if str(colC).upper() != "HMRC":
                                    monthDict[monthKey]["totalI"] += amountI
                                monthDict[monthKey]["totalH"] += amountH
                                currentRow += 1
                            # Output summary table headers
                            ws.cell(row=summaryRow, column=1).value = "Description"
                            ws.cell(row=summaryRow, column=2).value = "Liability"
                            ws.cell(row=summaryRow, column=3).value = "Payment"
                            ws.cell(row=summaryRow, column=4).value = "Difference"
                            for col in range(1, 5):
                                ws.cell(row=summaryRow, column=col).font = openpyxl.styles.Font(bold=True)
                            summaryRow += 1
                            # Opening balance row
                            ws.cell(row=summaryRow, column=1).value = "Opening Balance"
                            ws.cell(row=summaryRow, column=2).value = openingBalance
                            ws.cell(row=summaryRow, column=4).value = openingBalance
                            summaryRow += 1
                            # Output each month's calculated values
                            for monthKey in monthDict:
                                liabilities = monthDict[monthKey]["totalI"]
                                payment = monthDict[monthKey]["totalH"]
                                difference = liabilities - payment
                                ws.cell(row=summaryRow, column=1).value = monthKey
                                ws.cell(row=summaryRow, column=2).value = liabilities
                                ws.cell(row=summaryRow, column=3).value = payment
                                ws.cell(row=summaryRow, column=4).value = difference
                                totalLiability += liabilities
                                totalPayment += payment
                                summaryRow += 1
                            # Add opening balance to total liabilities
                            totalLiability += openingBalance
                            # Write the Balance row
                            totalOutstanding = totalLiability - totalPayment
                            ws.cell(row=summaryRow, column=1).value = "Balance"
                            ws.cell(row=summaryRow, column=2).value = totalLiability
                            ws.cell(row=summaryRow, column=3).value = totalPayment
                            ws.cell(row=summaryRow, column=4).value = totalOutstanding
                            found = True
                            break
                        currentRow += 1
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format7 ---
                elif fmt == 'format7':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    accountName = ws.cell(row=4, column=1).value
                    summaryRow = 15
                    lastRow = wsTrans.max_row
                    currentRow = 1
                    monthDict = {}
                    found = False
                    while currentRow <= lastRow:
                        if wsTrans.cell(row=currentRow, column=1).value == accountName:
                            openingBalance = wsTrans.cell(row=currentRow + 1, column=9).value or 0
                            currentRow = currentRow + 2
                            totalOutstanding = 0
                            # Process transactions until a 'Total' or 'Closing Balance' is encountered
                            while (currentRow <= lastRow and
                                wsTrans.cell(row=currentRow, column=1).value not in (None, "", "Closing Balance") and
                                not (str(wsTrans.cell(row=currentRow, column=1).value).startswith("Total"))):
                                amountI = wsTrans.cell(row=currentRow, column=9).value or 0
                                amountH = wsTrans.cell(row=currentRow, column=8).value or 0
                                transactionDate = wsTrans.cell(row=currentRow, column=1).value
                                # Extract monthKey
                                if isinstance(transactionDate, datetime):
                                    monthKey = transactionDate.strftime("%B %Y")
                                elif transactionDate:
                                    try:
                                        dt = pd.to_datetime(transactionDate, errors='coerce')
                                        if pd.notna(dt):
                                            monthKey = dt.strftime("%B %Y")
                                        else:
                                            monthKey = str(transactionDate)[:7]
                                    except Exception:
                                        monthKey = str(transactionDate)[:7]
                                else:
                                    monthKey = ""
                                if monthKey not in monthDict:
                                    monthDict[monthKey] = {"liability": 0, "payment": 0}
                                # Update monthly liabilities for 'Manual Journal' (Column B)
                                colB = wsTrans.cell(row=currentRow, column=2).value or ""
                                if str(colB) == "Manual Journal":
                                    monthDict[monthKey]["liability"] += amountI - amountH
                                # Add to payment if Column B is 'Spend Money'
                                if str(colB) == "Spend Money":
                                    monthDict[monthKey]["payment"] += amountH
                                currentRow += 1
                            # Output summary table headers
                            ws.cell(row=summaryRow, column=1).value = "Month"
                            ws.cell(row=summaryRow, column=2).value = "Liability"
                            ws.cell(row=summaryRow, column=3).value = "Payment"
                            ws.cell(row=summaryRow, column=4).value = "Outstanding"
                            for col in range(1, 5):
                                ws.cell(row=summaryRow, column=col).font = openpyxl.styles.Font(bold=True)
                            summaryRow += 1
                            # Opening balance row
                            ws.cell(row=summaryRow, column=1).value = "Opening Balance"
                            ws.cell(row=summaryRow, column=2).value = openingBalance
                            ws.cell(row=summaryRow, column=4).value = openingBalance
                            summaryRow += 1
                            # Output each month's calculated values
                            for monthKey in monthDict:
                                liabilities = monthDict[monthKey]["liability"]
                                payment = monthDict[monthKey]["payment"]
                                outstanding = liabilities - payment
                                ws.cell(row=summaryRow, column=1).value = monthKey
                                ws.cell(row=summaryRow, column=2).value = liabilities
                                ws.cell(row=summaryRow, column=3).value = payment
                                # Display Outstanding only if nonzero, else leave blank
                                if outstanding != 0:
                                    ws.cell(row=summaryRow, column=4).value = outstanding
                                else:
                                    ws.cell(row=summaryRow, column=4).value = None
                                summaryRow += 1
                            # Calculate the total outstanding from the monthly entries including opening balance
                            totalOutstanding = openingBalance
                            for monthKey in monthDict:
                                totalOutstanding += monthDict[monthKey]["liability"] - monthDict[monthKey]["payment"]
                            ws.cell(row=summaryRow, column=1).value = "Outstanding Total"
                            ws.cell(row=summaryRow, column=4).value = totalOutstanding
                            found = True
                            break
                        currentRow += 1
                    if not found:
                        print(f"Account not found for Format7: {accountName}")
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format8 ---
                elif fmt == 'format8':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    accountName = ws.cell(row=4, column=1).value
                    summaryRow = 15
                    lastRow = wsTrans.max_row
                    currentRow = 1
                    monthDict = {}
                    found = False
                    while currentRow <= lastRow:
                        if wsTrans.cell(row=currentRow, column=1).value == accountName:
                            # Opening balance from column 8 (H)
                            openingBalance = wsTrans.cell(row=currentRow + 1, column=8).value or 0
                            closingBalance = openingBalance
                            currentRow = currentRow + 2
                            # Process transactions until "Total PAYE" or "Closing Balance" is encountered
                            while (currentRow <= lastRow and
                                wsTrans.cell(row=currentRow, column=1).value not in (None, "", "Total PAYE", "Closing Balance")):
                                amountI = wsTrans.cell(row=currentRow, column=9).value or 0
                                amountH = wsTrans.cell(row=currentRow, column=8).value or 0
                                transactionDate = wsTrans.cell(row=currentRow, column=1).value
                                # Extract monthKey
                                if isinstance(transactionDate, datetime):
                                    monthKey = transactionDate.strftime("%B %Y")
                                elif transactionDate:
                                    try:
                                        dt = pd.to_datetime(transactionDate, errors='coerce')
                                        if pd.notna(dt):
                                            monthKey = dt.strftime("%B %Y")
                                        else:
                                            monthKey = str(transactionDate)[:7]
                                    except Exception:
                                        monthKey = str(transactionDate)[:7]
                                else:
                                    monthKey = ""
                                if monthKey not in monthDict:
                                    monthDict[monthKey] = {"receipts": 0, "payments": 0, "pdqDeposits": 0}
                                colB = str(wsTrans.cell(row=currentRow, column=2).value or "").upper()
                                # Receipts: 'RECEIVE MONEY' (Column H)
                                if colB == "RECEIVE MONEY":
                                    monthDict[monthKey]["receipts"] += amountH
                                # Payments: 'SPEND MONEY', 'PAYABLE PAYMENT', 'PAYABLE OVERPAYMENT' (Column I)
                                if colB in ("SPEND MONEY", "PAYABLE PAYMENT", "PAYABLE OVERPAYMENT"):
                                    monthDict[monthKey]["payments"] += amountI
                                # PDQ/Deposits: 'BANK TRANSFER' (I-H)
                                if colB == "BANK TRANSFER":
                                    monthDict[monthKey]["pdqDeposits"] += (amountI - amountH)
                                currentRow += 1
                            # Output summary table headers
                            ws.cell(row=summaryRow, column=1).value = "Month"
                            ws.cell(row=summaryRow, column=2).value = "Op Bal"
                            ws.cell(row=summaryRow, column=3).value = "Receipts"
                            ws.cell(row=summaryRow, column=4).value = "Payments"
                            ws.cell(row=summaryRow, column=5).value = "PDQ / Deposits"
                            ws.cell(row=summaryRow, column=6).value = "Clo Bal"
                            for col in range(1, 7):
                                ws.cell(row=summaryRow, column=col).font = openpyxl.styles.Font(bold=True)
                            summaryRow += 1
                            # Write the opening balance row
                            ws.cell(row=summaryRow, column=1).value = "Opening Balance"
                            ws.cell(row=summaryRow, column=2).value = openingBalance
                            summaryRow += 1
                            # Output each month's calculated values
                            for monthKey in monthDict:
                                receipts = monthDict[monthKey]["receipts"]
                                payments = monthDict[monthKey]["payments"]
                                pdqDeposits = monthDict[monthKey]["pdqDeposits"]
                                closingBalance = closingBalance + receipts - payments - pdqDeposits
                                ws.cell(row=summaryRow, column=1).value = monthKey
                                ws.cell(row=summaryRow, column=2).value = openingBalance
                                ws.cell(row=summaryRow, column=3).value = receipts
                                ws.cell(row=summaryRow, column=4).value = payments
                                ws.cell(row=summaryRow, column=5).value = pdqDeposits
                                ws.cell(row=summaryRow, column=6).value = closingBalance
                                # Update opening balance for next month
                                openingBalance = closingBalance
                                summaryRow += 1
                            found = True
                            break
                        currentRow += 1
                    if not found:
                        print(f"Account not found for Format8: {accountName}")
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format9 ---
                elif fmt == 'format9':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    # Get the period end date from cell A8
                    periodEndDate = ws.cell(row=8, column=1).value
                    # Validate date
                    try:
                        periodEndDate_dt = pd.to_datetime(periodEndDate, errors='raise')
                    except Exception:
                        print(f"Invalid date in A8 for Format9: {periodEndDate}")
                        continue
                    # Get the account head from A4
                    accountName = ws.cell(row=4, column=1).value
                    foundMatch = False
                    accountRow = None
                    for rowIndex in range(1, wsTrans.max_row + 1):
                        if wsTrans.cell(row=rowIndex, column=1).value == accountName:
                            foundMatch = True
                            accountRow = rowIndex
                            break
                    if not foundMatch:
                        print(f"Account head not found for Format9: {accountName}")
                        continue
                    # Move to the next row after the account header
                    accountRow = accountRow + 1
                    # Calculate the opening balance (H - I)
                    openingBalance = (wsTrans.cell(row=accountRow, column=8).value or 0) - (wsTrans.cell(row=accountRow, column=9).value or 0)
                    closingBalance = openingBalance
                    # Loop through transactions to calculate the closing balance
                    for transactionRow in range(accountRow + 1, wsTrans.max_row + 1):
                        val = wsTrans.cell(row=transactionRow, column=1).value
                        if val is not None and ("Total" in str(val) or "Closing Balance" in str(val)):
                            break
                        creditVal = wsTrans.cell(row=transactionRow, column=8).value or 0
                        debitVal = wsTrans.cell(row=transactionRow, column=9).value or 0
                        closingBalance = closingBalance + (creditVal - debitVal)
                    # Start populating the summary table from row 15
                    summaryStartRow = 15
                    ws.cell(row=summaryStartRow, column=1).value = "Date"
                    ws.cell(row=summaryStartRow, column=2).value = "Details"
                    ws.cell(row=summaryStartRow, column=3).value = "Amount £"
                    for col in range(1, 4):
                        ws.cell(row=summaryStartRow, column=col).font = openpyxl.styles.Font(bold=True)
                    summaryStartRow += 1
                    # Fill in the summary table with blank row
                    ws.cell(row=summaryStartRow, column=1).value = periodEndDate
                    ws.cell(row=summaryStartRow, column=2).value = "Balance as per statement"
                    ws.cell(row=summaryStartRow, column=3).value = ""
                    summaryStartRow += 1
                    # Add blank row
                    ws.cell(row=summaryStartRow, column=1).value = ""
                    ws.cell(row=summaryStartRow, column=2).value = ""
                    ws.cell(row=summaryStartRow, column=3).value = ""
                    summaryStartRow += 1
                    ws.cell(row=summaryStartRow, column=1).value = periodEndDate
                    ws.cell(row=summaryStartRow, column=2).value = "Balance per Control account"
                    ws.cell(row=summaryStartRow, column=3).value = closingBalance
                    ws.cell(row=summaryStartRow, column=3).number_format = "#,##0.00"
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)
                # --- Format10 ---
                elif fmt == 'format10':
                    ws = target_sheet
                    wsTrans = wb["Account Transactions"]
                    # Get the period end date from cell A8
                    periodEndDate = ws.cell(row=8, column=1).value
                    # Validate date
                    try:
                        periodEndDate_dt = pd.to_datetime(periodEndDate, errors='raise')
                    except Exception:
                        print(f"Invalid date in A8 for Format10: {periodEndDate}")
                        continue
                    # Get the account head from A4
                    accountName = ws.cell(row=4, column=1).value
                    foundMatch = False
                    accountRow = None
                    for rowIndex in range(1, wsTrans.max_row + 1):
                        if wsTrans.cell(row=rowIndex, column=1).value == accountName:
                            foundMatch = True
                            accountRow = rowIndex
                            break
                    if not foundMatch:
                        print(f"Account head not found for Format10: {accountName}")
                        continue
                    # Move to the next row after the account header
                    accountRow = accountRow + 1
                    # Calculate the opening balance (I - H)
                    openingBalance = (wsTrans.cell(row=accountRow, column=9).value or 0) - (wsTrans.cell(row=accountRow, column=8).value or 0)
                    closingBalance = openingBalance
                    # Loop through transactions to calculate the closing balance
                    for transactionRow in range(accountRow + 1, wsTrans.max_row + 1):
                        val = wsTrans.cell(row=transactionRow, column=1).value
                        if val is not None and ("Total" in str(val) or "Closing Balance" in str(val)):
                            break
                        closingBalance = closingBalance + ((wsTrans.cell(row=transactionRow, column=9).value or 0) - (wsTrans.cell(row=transactionRow, column=8).value or 0))
                    # Populate summary table
                    ws.cell(row=13, column=1).value = "Reconciliation"
                    ws.cell(row=13, column=1).font = openpyxl.styles.Font(bold=True, size=14)
                    ws.cell(row=15, column=1).value = "Date"
                    ws.cell(row=15, column=2).value = "£"
                    ws.cell(row=15, column=3).value = "Particular"
                    for col in range(1, 4):
                        ws.cell(row=15, column=col).font = openpyxl.styles.Font(bold=True)
                    ws.cell(row=16, column=1).value = periodEndDate
                    ws.cell(row=16, column=2).value = ""  # Manual input placeholder
                    ws.cell(row=16, column=3).value = "Balance as per "
                    ws.cell(row=17, column=1).value = periodEndDate
                    ws.cell(row=17, column=2).value = closingBalance
                    ws.cell(row=17, column=3).value = "Balance as per "
                    ws.cell(row=18, column=1).value = periodEndDate
                    # Set formula for difference (B16-B17)
                    ws.cell(row=18, column=2).value = f"=B16-B17"
                    ws.cell(row=18, column=3).value = "Difference"
                    # Apply accounting format to amount column (column 2)
                    for r in range(16, 19):
                        ws.cell(row=r, column=2).number_format = "#,##0.00"
                    # Apply summary table formatting
                    format_summary_table(ws, start_row=15)



    # --- Ensure FAR tables, depreciation rates, and months are built from the current workbook ---
    if 'FAR' in wb.sheetnames:
        ws_far = wb['FAR']
        # Remove gridlines from FAR sheet
        ws_far.sheet_view.showGridLines = False
        # Extract FAR data from the current workbook
        df_raw = pd.read_excel(excel_for_pandas, sheet_name='FAR', header=None, engine='openpyxl')
        static_headers = [
            'Purchase Date',
            'Details',
            'Cost',
            'Addition',
            'Total Cost',
            'Depreciation Rate',
            'Accumulated Depreciation'
        ]
        # Generate fiscal year month columns as strings like "Dep    Aug-24"
        months = []
        cur = fy_start.replace(day=1)
        while cur <= fy_end:
            months.append(f"Dep {cur.strftime('%b-%y')}")
            cur += pd.DateOffset(months=1)
        tables = []
        far_table_names = []
        far_table_deprates = {}
        i = 5  # Row 6 in Excel (0-based index)
        while i < len(df_raw):
            table_name = str(df_raw.iloc[i, 0])
            if table_name == 'nan' or not table_name.strip():
                i += 1
                continue
            far_table_names.append(table_name.strip())
            depreciation_row = i + 1
            dep_rate_val = None
            if depreciation_row < len(df_raw):
                dep_line = str(df_raw.iloc[depreciation_row, 0])
               
                match = re.search(r"Depreciation rate\s*:\s*([\d.]+)%", dep_line, re.IGNORECASE)
                if match:
                    dep_rate_val = float(match.group(1))
            if dep_rate_val is not None:
                far_table_deprates[table_name.strip()] = dep_rate_val
            header_row = i + 2
            data_start = i + 3
            data_end = data_start
            while data_end < len(df_raw):
                if str(df_raw.iloc[data_end, 1]).strip().lower() == 'total':
                    break
                data_end += 1
            table_data = df_raw.iloc[data_start:data_end, :len(static_headers)]
            table_data.columns = static_headers
            table_data = table_data.reset_index(drop=True)
            for m in months:
                table_data[m] = 0
            tables.append((table_name, table_data))
            i = data_end + 1
            while i < len(df_raw) and (str(df_raw.iloc[i, 0]) == 'nan' or not str(df_raw.iloc[i, 0]).strip()):
                i += 1

        # --- Merge new transactions and recalculate FAR tables before writing ---
        # Read Account Transactions sheet from the same Excel as FAR
        if 'Account Transactions' in wb.sheetnames:
            ws_trans = wb['Account Transactions']
            # Read as DataFrame for easier processing
            df_trans_raw = pd.read_excel(excel_for_pandas, sheet_name='Account Transactions', header=None, engine='openpyxl')
            header_row_idx = 4  # Row 5 in Excel (0-based)
            headers = list(df_trans_raw.iloc[header_row_idx].fillna('').astype(str))
            header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
            transactions = []
            i = header_row_idx + 1
            while i < len(df_trans_raw):
                account_name = str(df_trans_raw.iloc[i, 0])
                if account_name == 'nan' or not account_name.strip() or account_name.strip().lower() == 'total':
                    i += 1
                    continue
                if account_name not in far_table_names:
                    i += 1
                    continue
                data_start = i + 1
                data_end = data_start
                while data_end < len(df_trans_raw):
                    cell_val = str(df_trans_raw.iloc[data_end, 0]).strip().lower()
                    if cell_val.startswith('total'):
                        break
                    data_end += 1
                block = df_trans_raw.iloc[data_start:data_end, :].copy()
                block = block[~block.iloc[:,0].astype(str).str.strip().str.lower().isin(['opening balance','closing balance'])]
                if not block.empty:
                    mapped = pd.DataFrame()
                    mapped['Purchase Date'] = block.iloc[:, header_map['purchase date']] if 'purchase date' in header_map else block.iloc[:, 0]
                    mapped['Details'] = block.iloc[:, header_map['details']] if 'details' in header_map else block.iloc[:, 2]
                    mapped['Cost'] = block.iloc[:, 7] if block.shape[1] > 7 else 0
                    mapped['Asset Type'] = account_name.strip()
                    transactions.append(mapped)
                i = data_end + 1
            if transactions:
                df_trans = pd.concat(transactions, ignore_index=True)
            else:
                df_trans = pd.DataFrame()
        else:
            df_trans = pd.DataFrame()

        updated_tables = []
        for table_name, table_df in tables:
            updated_table = table_df.copy()
            if 'Details' in updated_table.columns:
                updated_table = updated_table[~updated_table['Details'].astype(str).str.strip().str.lower().isin(['opening balance','closing balance'])]
            else:
                if updated_table.shape[1] > 1:
                    updated_table = updated_table[~updated_table.iloc[:,1].astype(str).str.strip().str.lower().isin(['opening balance','closing balance'])]
            if 'Total Depreciation' not in updated_table.columns:
                updated_table['Total Depreciation'] = 0
            if 'WDV' not in updated_table.columns:
                updated_table['WDV'] = 0
            if not df_trans.empty and 'Asset Type' in df_trans.columns:
                matching_rows = df_trans[df_trans['Asset Type'] == table_name].reset_index(drop=True)
                if not matching_rows.empty:
                    new_rows = pd.DataFrame()
                    new_rows['Purchase Date'] = matching_rows['Purchase Date']
                    new_rows['Details'] = matching_rows['Details']
                    new_rows['Cost'] = matching_rows['Cost']
                    new_rows['Addition'] = 0
                    new_rows['Total Cost'] = 0
                    new_rows['Depreciation Rate'] = far_table_deprates.get(table_name.strip(), 0)
                    new_rows['Accumulated Depreciation'] = 0
                    for m in months:
                        new_rows[m] = 0
                    if 'Details' in new_rows.columns:
                        new_rows = new_rows[~new_rows['Details'].astype(str).str.strip().str.lower().isin(['opening balance','closing balance'])]
                    if not updated_table.empty:
                        merge_cols = ['Purchase Date', 'Details', 'Cost']
                        def safe_str(x):
                            if pd.isnull(x):
                                return ''
                            if isinstance(x, pd.Timestamp) or isinstance(x, datetime):
                                return x.strftime('%Y-%m-%d')
                            try:
                                dt = pd.to_datetime(x, errors='coerce')
                                if pd.notnull(dt):
                                    return dt.strftime('%Y-%m-%d')
                            except Exception:
                                pass
                            return str(x)
                        updated_table_key = updated_table[merge_cols].applymap(safe_str).agg('|'.join, axis=1)
                        new_rows_key = new_rows[merge_cols].applymap(safe_str).agg('|'.join, axis=1)
                        new_rows = new_rows[~new_rows_key.isin(updated_table_key)]
                    for idx, row in new_rows.iterrows():
                        try:
                            pdate = pd.to_datetime(row['Purchase Date'], errors='coerce')
                        except Exception:
                            pdate = None
                        purchase_amt = pd.to_numeric(row['Cost'], errors='coerce') if pd.notnull(row['Cost']) else 0
                        if pd.notnull(pdate) and fy_start <= pdate <= fy_end:
                            new_rows.at[idx, 'Addition'] = purchase_amt
                            new_rows.at[idx, 'Cost'] = 0
                        else:
                            new_rows.at[idx, 'Addition'] = 0
                            new_rows.at[idx, 'Cost'] = purchase_amt
                    new_rows['Total Cost'] = pd.to_numeric(new_rows['Cost'], errors='coerce').fillna(0) + pd.to_numeric(new_rows['Addition'], errors='coerce').fillna(0)
                    for idx, row in new_rows.iterrows():
                        try:
                            pdate = pd.to_datetime(row['Purchase Date'], errors='coerce')
                        except Exception:
                            pdate = None
                        total_cost = pd.to_numeric(row['Total Cost'], errors='coerce')
                        dep_rate = pd.to_numeric(row['Depreciation Rate'], errors='coerce')
                        last_fy_month = fy_start - pd.DateOffset(months=1)
                        months_in_use = 0
                        if pd.notnull(pdate) and pdate <= last_fy_month:
                            months_in_use = (last_fy_month.year - pdate.year) * 12 + (last_fy_month.month - pdate.month) + 1
                            if months_in_use < 0:
                                months_in_use = 0
                        monthly_dep = (total_cost * dep_rate / 100) / 12 if total_cost and dep_rate else 0
                        acc_dep = monthly_dep * months_in_use
                        new_rows.at[idx, 'Accumulated Depreciation'] = acc_dep
                    for idx, row in new_rows.iterrows():
                        try:
                            pdate = pd.to_datetime(row['Purchase Date'], errors='coerce')
                        except Exception:
                            pdate = None
                        total_cost = pd.to_numeric(row['Total Cost'], errors='coerce')
                        dep_rate = pd.to_numeric(row['Depreciation Rate'], errors='coerce')
                        monthly_dep = (total_cost * dep_rate / 100) / 12 if total_cost and dep_rate else 0
                        for m in months:
                            try:
                                m_str = m.replace('Dep', '').strip()
                                m_dt = pd.to_datetime(m_str, format='%b-%y')
                            except Exception:
                                m_dt = None
                            if pd.notnull(pdate) and pd.notnull(m_dt):
                                if (pdate <= m_dt + pd.offsets.MonthEnd(0)) and (m_dt <= mgmt_acct_month):
                                    new_rows.at[idx, m] = monthly_dep
                                elif m_dt > mgmt_acct_month:
                                    new_rows.at[idx, m] = ''
                                else:
                                    new_rows.at[idx, m] = 0
                            else:
                                new_rows.at[idx, m] = ''
                    dep_month_cols = months
                    acc_dep_col_name = 'Accumulated Depreciation'
                    total_cost_col_name = 'Total Cost'
                    new_rows['Total Depreciation'] = pd.to_numeric(new_rows[acc_dep_col_name], errors='coerce').fillna(0) + new_rows[dep_month_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
                    new_rows['WDV'] = pd.to_numeric(new_rows[total_cost_col_name], errors='coerce').fillna(0) - new_rows['Total Depreciation']
                    if 'Total Depreciation' not in updated_table.columns:
                        updated_table['Total Depreciation'] = 0
                    if 'WDV' not in updated_table.columns:
                        updated_table['WDV'] = 0
                    full_cols = static_headers + months + ['Total Depreciation', 'WDV']
                    updated_table = updated_table.reindex(columns=full_cols, fill_value=0)
                    new_rows = new_rows.reindex(columns=full_cols, fill_value=0)
                    if not new_rows.empty:
                        print(f"Appending {len(new_rows)} new transaction row(s) to table '{table_name}'")
                    combined_table = pd.concat([updated_table, new_rows], ignore_index=True)
                    updated_table = combined_table
            dep_rate_val = far_table_deprates.get(table_name.strip(), 0)
            for col in ['Addition', 'Total Cost', 'Depreciation Rate', 'Accumulated Depreciation'] + months:
                if col not in updated_table.columns:
                    updated_table[col] = 0
            updated_table['Depreciation Rate'] = dep_rate_val
            for idx, row in updated_table.iterrows():
                try:
                    pdate = pd.to_datetime(row['Purchase Date'], errors='coerce')
                except Exception:
                    pdate = None
                cost_val = pd.to_numeric(row['Cost'], errors='coerce') if pd.notnull(row['Cost']) else 0
                addition_val = pd.to_numeric(row['Addition'], errors='coerce') if pd.notnull(row['Addition']) else 0
                orig_amt = cost_val + addition_val
                if pd.notnull(pdate) and fy_start <= pdate <= fy_end:
                    updated_table.at[idx, 'Addition'] = orig_amt
                    updated_table.at[idx, 'Cost'] = 0
                else:
                    updated_table.at[idx, 'Addition'] = 0
                    updated_table.at[idx, 'Cost'] = orig_amt
            updated_table['Total Cost'] = pd.to_numeric(updated_table['Cost'], errors='coerce').fillna(0) + pd.to_numeric(updated_table['Addition'], errors='coerce').fillna(0)
            for idx, row in updated_table.iterrows():
                try:
                    pdate = pd.to_datetime(row['Purchase Date'], errors='coerce')
                except Exception:
                    pdate = None
                total_cost = pd.to_numeric(row['Total Cost'], errors='coerce')
                dep_rate = pd.to_numeric(row['Depreciation Rate'], errors='coerce')
                last_fy_month = fy_start - pd.DateOffset(months=1)
                months_in_use = 0
                if pd.notnull(pdate) and pdate <= last_fy_month:
                    months_since_purchase = (last_fy_month.year - pdate.year) * 12 + (last_fy_month.month - pdate.month) + 1
                    if months_since_purchase < 0:
                        months_since_purchase = 0
                    monthly_dep = (total_cost * dep_rate / 100) / 12 if total_cost and dep_rate else 0
                    months_to_fully_depreciate = int(total_cost // monthly_dep) if monthly_dep > 0 else 0
                    months_in_use = min(months_since_purchase, months_to_fully_depreciate)
                    acc_dep = monthly_dep * months_in_use
                    if months_since_purchase >= months_to_fully_depreciate:
                        acc_dep = total_cost
                    updated_table.at[idx, 'Accumulated Depreciation'] = acc_dep
                else:
                    monthly_dep = (total_cost * dep_rate / 100) / 12 if total_cost and dep_rate else 0
                    updated_table.at[idx, 'Accumulated Depreciation'] = 0
            for idx2, row2 in updated_table.iterrows():
                try:
                    pdate2 = pd.to_datetime(row2['Purchase Date'], errors='coerce')
                except Exception:
                    pdate2 = None
                total_cost2 = pd.to_numeric(row2['Total Cost'], errors='coerce')
                dep_rate2 = pd.to_numeric(row2['Depreciation Rate'], errors='coerce')
                monthly_dep2 = (total_cost2 * dep_rate2 / 100) / 12 if total_cost2 and dep_rate2 else 0
                for m in months:
                    updated_table.iloc[idx2, updated_table.columns.get_loc(m)] = 0
                total_dep_so_far2 = pd.to_numeric(updated_table.at[idx2, 'Accumulated Depreciation'], errors='coerce') if 'Accumulated Depreciation' in updated_table.columns else 0
                for m in months:
                    try:
                        m_str = m.replace('Dep', '').strip()
                        m_dt = pd.to_datetime(m_str, format='%b-%y')
                    except Exception:
                        m_dt = None
                    if pd.notnull(pdate2) and pd.notnull(m_dt):
                        if (pdate2 <= m_dt + pd.offsets.MonthEnd(0)) and (m_dt <= mgmt_acct_month):
                            if total_dep_so_far2 < total_cost2:
                                dep_this_month2 = min(monthly_dep2, total_cost2 - total_dep_so_far2)
                                updated_table.iloc[idx2, updated_table.columns.get_loc(m)] = dep_this_month2
                                total_dep_so_far2 += dep_this_month2
                            else:
                                updated_table.iloc[idx2, updated_table.columns.get_loc(m)] = 0
                        elif m_dt > mgmt_acct_month:
                            updated_table.iloc[idx2, updated_table.columns.get_loc(m)] = ''
                        else:
                            updated_table.iloc[idx2, updated_table.columns.get_loc(m)] = 0
                    else:
                        updated_table.iloc[idx2, updated_table.columns.get_loc(m)] = ''
            dep_month_cols = months
            acc_dep_col_name = 'Accumulated Depreciation'
            total_cost_col_name = 'Total Cost'
            updated_table['Total Depreciation'] = pd.to_numeric(updated_table[acc_dep_col_name], errors='coerce').fillna(0) + updated_table[dep_month_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
            updated_table['WDV'] = pd.to_numeric(updated_table[total_cost_col_name], errors='coerce').fillna(0) - updated_table['Total Depreciation']
            final_cols = static_headers + months + ['Total Depreciation', 'WDV']
            updated_table = updated_table.reindex(columns=final_cols, fill_value=0)
            updated_tables.append((table_name, updated_table.copy()))

        # Now write all updated_tables to the consolidated output (as before)
        from openpyxl.styles.borders import Border, Side
        from openpyxl.styles import Font, Alignment
        for row in ws_far.iter_rows(min_row=6, max_row=ws_far.max_row):
            for cell in row:
                cell.value = None
                cell.border = Border()
        write_row = 6
        if updated_tables:
            for table_name, combined_df in updated_tables:
                combined_df = combined_df.dropna(how='all').reset_index(drop=True)
                combined_df = combined_df.applymap(lambda x: x.item() if hasattr(x, 'item') else x)
                header_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                uniform_font = Font(name='Gill Sans MT', size=12, bold=False)
                bold_font = Font(name='Gill Sans MT', size=12, bold=True)
                center_align = Alignment(horizontal='center', vertical='center')
                right_align = Alignment(horizontal='right', vertical='center')
                left_align = Alignment(horizontal='left', vertical='center')
                date_align = Alignment(horizontal='center', vertical='center')
                table_name_cell = ws_far.cell(row=write_row, column=1, value=table_name)
                table_name_cell.font = bold_font
                for col in range(2, len(combined_df.columns)+1):
                    ws_far.cell(row=write_row, column=col).font = uniform_font
                    ws_far.cell(row=write_row, column=col).border = Border()
                dep_rate_val = None
                for tname, rate in far_table_deprates.items():
                    if tname == table_name.strip():
                        dep_rate_val = rate
                        break
                dep_line = f"Depreciation rate: {dep_rate_val:.0f}%" if dep_rate_val is not None else "Depreciation rate: "
                dep_line_cell = ws_far.cell(row=write_row+1, column=1, value=dep_line)
                dep_line_cell.font = uniform_font
                dep_line_cell.alignment = left_align
                for col in range(2, len(combined_df.columns)+1):
                    ws_far.cell(row=write_row+1, column=col).font = uniform_font
                    ws_far.cell(row=write_row+1, column=col).border = Border()
                mgmt_acct_end_str = ''
                if mgmt_acct_month is not None:
                    mgmt_acct_end_str = mgmt_acct_month.strftime('%b-%Y')
                col_headers = list(combined_df.columns)
                if len(col_headers) >= 2:
                    col_headers[-2] = f"{col_headers[-2]} (as at {mgmt_acct_end_str})"
                    col_headers[-1] = f"{col_headers[-1]} (as at {mgmt_acct_end_str})"
                for col_idx, col_name in enumerate(col_headers, start=1):
                    cell = ws_far.cell(row=write_row+2, column=col_idx, value=col_name)
                    cell.border = header_border
                    cell.font = bold_font
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                units_row_idx = write_row+3
                units_map = {}
                for col in combined_df.columns:
                    col_lc = str(col).strip().lower()
                    if 'date' in col_lc:
                        units_map[col] = ''
                    elif 'rate' in col_lc:
                        units_map[col] = '%'
                    elif 'depreciation' in col_lc or 'cost' in col_lc or 'addition' in col_lc or 'wdv' in col_lc:
                        units_map[col] = '£'
                    elif 'dep' in col_lc and any(mon in col_lc for mon in ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']):
                        units_map[col] = '£'
                    else:
                        units_map[col] = ''
                for col_idx, col_name in enumerate(combined_df.columns, start=1):
                    cell = ws_far.cell(row=units_row_idx, column=col_idx, value=units_map.get(col_name, ''))
                    cell.font = uniform_font
                    cell.alignment = center_align
                    cell.border = header_border
                # Write data rows, and apply row-wise SUM formulas for Total Depreciation and WDV
                dep_col_idx = None
                wdv_col_idx = None
                acc_dep_col_idx = None
                total_cost_col_idx = None
                # Find column indices for relevant columns
                for idx, col_name in enumerate(combined_df.columns):
                    col_clean = str(col_name).strip().lower().replace(' ', '')
                    if col_clean == 'totaldepreciation':
                        dep_col_idx = idx
                    if col_clean == 'wdv':
                        wdv_col_idx = idx
                    if col_clean == 'accumulateddepreciation':
                        acc_dep_col_idx = idx
                    if col_clean == 'totalcost':
                        total_cost_col_idx = idx
                month_col_indices = [i for i, col in enumerate(combined_df.columns) if any(mon in str(col).lower() for mon in ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'])]
                for r in range(len(combined_df)):
                    for c in range(len(combined_df.columns)):
                        val = combined_df.iloc[r, c]
                        if hasattr(val, 'item'):
                            val = val.item()
                        cell = ws_far.cell(row=write_row+4+r, column=c+1)
                        col_name = combined_df.columns[c]
                        col_lc = str(col_name).strip().lower()
                        # Row-wise SUM for Total Depreciation
                        if dep_col_idx is not None and c == dep_col_idx and acc_dep_col_idx is not None and month_col_indices:
                            acc_dep_cell = ws_far.cell(row=write_row+4+r, column=acc_dep_col_idx+1).coordinate
                            month_cells = [ws_far.cell(row=write_row+4+r, column=idx+1).coordinate for idx in month_col_indices]
                            cell.value = f"=SUM({acc_dep_cell},{','.join(month_cells)})"
                        # Row-wise SUM for WDV
                        elif wdv_col_idx is not None and c == wdv_col_idx and total_cost_col_idx is not None and dep_col_idx is not None:
                            total_cost_cell = ws_far.cell(row=write_row+4+r, column=total_cost_col_idx+1).coordinate
                            dep_cell = ws_far.cell(row=write_row+4+r, column=dep_col_idx+1).coordinate
                            cell.value = f"={total_cost_cell}-{dep_cell}"
                        else:
                            cell.value = val
                        # Add vertical borders to data cells
                        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'))
                        cell.font = uniform_font
                        if c == 0:
                            cell.alignment = date_align
                            if pd.notnull(val):
                                try:
                                    cell.number_format = 'mm-dd-yyyy'
                                except Exception:
                                    pass
                        elif 'rate' in col_lc:
                            cell.alignment = right_align
                            try:
                                if pd.notnull(val):
                                    cell.value = f"{float(val):.0f}%"
                                    cell.number_format = '0%'
                            except Exception:
                                pass
                        elif isinstance(val, (int, float)):
                            cell.alignment = right_align
                            try:
                                if pd.notnull(val):
                                    cell.number_format = '#,##0.00_);(#,##0.00);"-"??_);_(@_)'
                            except Exception:
                                pass
                        elif 'details' in col_lc:
                            cell.alignment = left_align
                        else:
                            cell.alignment = left_align
                total_row = [''] * len(combined_df.columns)
                total_row[1] = 'Total'
                data_for_total = combined_df.copy()
                data_for_total = data_for_total[data_for_total.iloc[:,1].astype(str).str.lower() != 'total']
                data_for_total = data_for_total.dropna(how='all')
                for idx, col in enumerate(combined_df.columns):
                    col_lc = str(col).strip().lower()
                    if 'rate' in col_lc:
                        total_row[idx] = ''
                        continue
                    try:
                        col_data = data_for_total[col]
                        if isinstance(col_data, pd.DataFrame):
                            col_data = col_data.iloc[:,0]
                        if isinstance(col_data, (pd.Series, list, tuple)):
                            col_numeric = pd.to_numeric(col_data, errors='coerce')
                            if pd.api.types.is_numeric_dtype(col_numeric):
                                total_row[idx] = col_numeric.sum(skipna=True)
                    except Exception as e:
                        pass
                total_row_idx = write_row + 4 + len(combined_df) - 1 + 1
                for c in range(len(total_row)):
                    cell = ws_far.cell(row=total_row_idx, column=c+1)
                    col_name = str(combined_df.columns[c]) if c < len(combined_df.columns) else ''
                    col_name_clean = col_name.strip().lower().replace(' ', '')
                    if c == 0:
                        cell.value = None
                        cell.alignment = date_align
                    elif c == 1:
                        cell.value = 'Total'
                        cell.alignment = left_align
                    elif 'rate' in col_name_clean:
                        cell.value = ''
                        cell.alignment = right_align
                    elif col_name_clean in ['totaldepreciation', 'wdv']:
                        col_letter = openpyxl.utils.get_column_letter(c+1)
                        data_start_row = write_row + 4
                        data_end_row = data_start_row + len(combined_df) - 1
                        cell.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})"
                        cell.alignment = right_align
                    else:
                        col_letter = openpyxl.utils.get_column_letter(c+1)
                        data_start_row = write_row + 4
                        data_end_row = data_start_row + len(combined_df) - 1
                        cell.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})"
                        cell.alignment = right_align
                    cell.border = header_border
                    cell.font = bold_font
                    if c == 0:
                        try:
                            cell.number_format = 'mm-dd-yyyy'
                        except Exception:
                            pass
                    elif c > 0:
                        try:
                            if not ('rate' in col_name_clean or c == 1):
                                cell.number_format = '#,##0.00_);(#,##0.00);"-"??_);_(@_)'
                        except Exception:
                            pass
                # Auto-adjust column widths for this table in FAR
                for col_idx, col_name in enumerate(combined_df.columns, start=1):
                    max_length = len(str(col_name))
                    for row_offset in range(0, len(combined_df) + 4):
                        cell = ws_far.cell(row=write_row + 2 + row_offset, column=col_idx)
                        val = cell.value
                        if val is not None:
                            val_str = str(val)
                            if len(val_str) > max_length:
                                max_length = len(val_str)
                    ws_far.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = max_length + 2
                write_row = total_row_idx+3

        # --- Auto-adjust column widths for all tables in all sheets (except excluded) ---
        excluded_sheets = ['FAR', 'Mappings', 'Account Transactions', 'Sheet1', 'Sheet2', 'Sheet3']
        for ws in wb.worksheets:
            if ws.title in excluded_sheets:
                continue
            # Find the header row (assume always at row 15 as per your logic)
            header_row = 15
            # Find the number of columns by checking for non-empty cells in header
            max_col = 0
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=header_row, column=col).value:
                    max_col = col
            if max_col == 0:
                continue
            # For each column, find the max length among header, units, and all data rows
            for col_idx in range(1, max_col + 1):
                max_length = 0
                # Header
                val = ws.cell(row=header_row, column=col_idx).value
                if val is not None:
                    max_length = len(str(val))
                # Units row
                val = ws.cell(row=header_row + 1, column=col_idx).value
                if val is not None and len(str(val)) > max_length:
                    max_length = len(str(val))
                # Data rows (assume data starts at header_row+2 and goes until first empty in col 2)
                data_row = header_row + 2
                while True:
                    val = ws.cell(row=data_row, column=2).value
                    if val is None or str(val).strip() == '':
                        break
                    val = ws.cell(row=data_row, column=col_idx).value
                    if val is not None and len(str(val)) > max_length:
                        max_length = len(str(val))
                    data_row += 1
                ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = max_length + 2





    # Apply formatting and adjust column widths after each FormatN routine
    from openpyxl.styles import PatternFill
    excludeSheets = ["Account Transactions", "P&L", "Corporation Tax", "Mappings", "FAR"]
    for ws in wb.worksheets:
        if ws.title not in excludeSheets:
            # Find the last row in the summary table (starting from A15)
            max_row = ws.max_row
            lastRow = None
            for r in range(max_row, 14, -1):
                val = ws.cell(row=r, column=1).value
                if val is not None and str(val).strip() != "":
                    lastRow = r
                    break
            if lastRow is not None and lastRow >= 15:
                # Find the last used column in lastRow
                max_col = ws.max_column
                lastCol = None
                for c in range(max_col, 0, -1):
                    val = ws.cell(row=lastRow, column=c).value
                    if val is not None and str(val).strip() != "":
                        lastCol = c
                        break
                summaryValue = None
                rng = None
                # Loop through columns in lastRow to find the last numeric value
                for i in range(1, (lastCol or 0) + 1):
                    cell = ws.cell(row=lastRow, column=i)
                    val = cell.value
                    if val is not None and isinstance(val, (int, float)) and str(val).strip() != "":
                        summaryValue = val
                        rng = cell
                # Check if C8 is empty before updating it
                c8 = ws.cell(row=8, column=3)
                if (c8.value is None or str(c8.value).strip() == ""):
                    if summaryValue is not None:
                        # Highlight the summary value cell
                        if rng is not None:
                            rng.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                        # Copy the summary value to C8 and format conditionally
                        c8.value = summaryValue
                        if summaryValue < 0:
                            c8.number_format = "#,##0.00;(#,##0.00)"
                        elif summaryValue == 0:
                            c8.number_format = "#,##0.00;-#,##0.00;-"
                        else:
                            c8.number_format = "#,##0.00;(#,##0.00)"
                    else:
                        c8.value = "No numeric summary value"
                # --- Adjust column widths for summary table ---
                for col_idx in range(1, (lastCol or 0) + 1):
                    max_length = 0
                    # Check header (row 15)
                    val = ws.cell(row=15, column=col_idx).value
                    if val is not None:
                        max_length = len(str(val))
                    # Check units row (row 16)
                    val = ws.cell(row=16, column=col_idx).value
                    if val is not None and len(str(val)) > max_length:
                        max_length = len(str(val))
                    # Check all data rows
                    for data_row in range(17, lastRow + 1):
                        val = ws.cell(row=data_row, column=col_idx).value
                        if val is not None and len(str(val)) > max_length:
                            max_length = len(str(val))
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = max_length + 2
            else:
                ws.cell(row=8, column=3).value = "No data found"

    # --- MODULE 5: DeleteSheetsWithNoData ---
    ws_to_delete = []
    for ws in wb.worksheets:
            c8 = ws.cell(row=8, column=3).value
            if c8 == "No data found":
                ws_to_delete.append(ws.title)
    for sheet_name in ws_to_delete:
            del wb[sheet_name]
    if ws_to_delete:
            print(f"Deleted sheets with 'No data found' in C8: {', '.join(ws_to_delete)}")

        # 4. Save the modified workbook to output folder
    safe_name = re.sub(r'[^A-Za-z0-9._-]', '_', str(output_base_name))
    far_output_name = f'{safe_name}.xlsx'
    output_path = os.path.join(output_folder, far_output_name)
    wb.save(output_path)
    print(f"Done! Output saved to: {output_path}")


    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

import streamlit as st

# --- Custom CSS for full-width layout and bigger logo ---
st.markdown("""
    <style>
    .main, .stApp {
        background-color: #f7f9fa;
        width: 100%;
        min-height: 100vh;
        margin: 0;
        padding: 0;
    }
    .header-bar {
    background: linear-gradient(90deg, #003366 0%, #005fa3 100%);
    color: white;
    padding: 1.2rem 0 0.7rem 0;
    margin-bottom: 1.5rem;
    border-radius: 0 0 18px 18px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    width: 100%;         /* Full viewport width */
    margin-left: 0;
    margin-right: 0;
    box-sizing: border-box;
}
    .card {
        background: white;
        border-radius: 16px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        padding: 2rem 2.5rem 2rem 2.5rem;
        margin: 2rem auto;
        max-width: 1200px; /* increase width */
        width: 95vw; /* fill viewport */
    }
    .centered {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    .top-left-logo {
        position: fixed;
        top: 18px;
        left: 18px;
        z-index: 1000;
        background: transparent;
    }
    </style>
""", unsafe_allow_html=True)


# --- Header Bar with Logo and Title ---
st.markdown('<div class="top-left-logo">', unsafe_allow_html=True)
st.image("Corient_Logo-01.jpg", width=160)  # Bigger logo
st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
    <div class="header-bar">
        <h1 style="margin-bottom:0.2rem; font-family:Gill Sans MT,Arial,sans-serif; font-size:2.5rem;">INNControl Processor</h1>
        
    </div>
""", unsafe_allow_html=True)


# --- Card-style Upload Section ---
with st.container():
    upload_card = """
    <style>
    .card.centered {background: white; border-radius: 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); padding: 2rem 2.5rem 2rem 2.5rem; margin: 0 auto 2rem auto; max-width: 500px; display: flex; flex-direction: column; align-items: center;}
    </style>
    <div class="card centered">
    """
    st.markdown(upload_card, unsafe_allow_html=True)


from io import BytesIO
from openpyxl import load_workbook
import pandas as pd
import streamlit as st

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    # ✅ Read file once
    file_content = uploaded_file.read()

    # ✅ Reuse the content for both libraries
    excel_for_openpyxl = BytesIO(file_content)
    excel_for_pandas = BytesIO(file_content)

    # ✅ Load workbook and data
    try:
        wb = load_workbook(excel_for_openpyxl)
        xls = pd.ExcelFile(excel_for_pandas, engine="openpyxl")

       # for sheet in xls.sheet_names:
        #    df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
         #   st.write(f"Sheet: {sheet}")
         #   st.dataframe(df)

    except Exception as e:
        st.error(f"❌ Processing failed: {str(e)}")
else:
    st.warning("Please upload an excel file to continue.")



    st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is not None:
    with st.container():
        process_card = """
        <style>
        .card.centered {background: white; border-radius: 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); padding: 2rem 2.5rem 2rem 2.5rem; margin: 0 auto 2rem auto; max-width: 500px; display: flex; flex-direction: column; align-items: center;}
        </style>
        <div class="card centered">
        """
        st.markdown(process_card, unsafe_allow_html=True)
        if 'processing' not in st.session_state:
            st.session_state['processing'] = False
        if not st.session_state['processing']:
            if st.button('Start Processing', key='start_processing'):
                st.session_state['processing'] = True
                st.rerun()
        else:
            st.markdown(f"<div style='font-size:1.1rem;'><span style='font-size:1.5rem;'>📄</span> <b>Processing:</b> {uploaded_file.name}</div>", unsafe_allow_html=True)
            with st.spinner("Processing... Please wait."):
                st.image("https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExY2Uwbm5xeTQwd3A4eWU2bGU4Ym9pdjF0YWRkZzltcGJyOTE5czJ0ZCZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/2bYewTk7K2No1NvcuK/giphy.gif", width=200)
                try:
                   
                    output_bytes = process_far_file(file_content)
                    output_data = output_bytes.getvalue()
                    st.session_state['processing'] = False
                    st.markdown("<div style='font-size:1.1rem; color: #155724;'><span style='font-size:1.5rem;'>✅</span> <b>File processed successfully!</b> Download is ready.</div>", unsafe_allow_html=True)
                    st.download_button(
                        label="⬇️ Download Processed Excel",
                        data=output_data,
                        file_name="MA_Processed.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.session_state['processing'] = False
                    st.markdown(f"<div style='font-size:1.1rem; color: #721c24;'><span style='font-size:1.5rem;'>❌</span> <b>Processing failed:</b> {e}</div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
