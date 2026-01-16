"""Report import utilities translated from the original VBA Module3.

This module handles importing data from other report workbooks, including
cross-version migrations, sheet-by-sheet transfers, and OpenOffice handling.
"""

from __future__ import annotations

from typing import Any, Iterable, List, Optional

from .constants import (
    VB_DEFAULT_BUTTON1,
    VB_EXCLAMATION,
    VB_OK,
    VB_OK_CANCEL,
    VB_OK_ONLY,
    XL_AUTOMATIC,
    XL_MANUAL,
)
from .helpers import in_array as in_array_helper, msg_box

PWORD = "SCoE"
LPWORD = "KCoE"


tgtWB: Optional[Any] = None
srcWB: Optional[Any] = None
isOO = False


def importFromReport(app: Any, workbooks: Any, active_workbook: Any) -> None:
    """Import data from another report workbook into the active workbook."""
    global tgtWB, srcWB, isOO

    # Capture open workbook names so the new import workbook can be identified.
    open_wbs: List[str] = ["" for _ in range(workbooks.Count + 1)]
    idx = 1
    for wb in workbooks:
        open_wbs[idx] = wb.Name
        idx += 1

    isOO = False

    # PayPal forms do not support report imports.
    if active_workbook.Sheets("Contents").Range("B39") == "PAYPAL":
        msg_box("This doesn't apply to PAYPAL form", VB_OK_ONLY, "")
        return

    msg = (
        "You are about to import from a different report form! This will overwrite ALL UNSAVED data already in this workbook."
    )
    msg += "\n\nThe report will be saved in a new file based on the imported report's branch name."
    title = "IMPORT Report"
    style = VB_OK_CANCEL + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    if msg_box(msg, style, title) != VB_OK:
        return

    app.DisplayStatusBar = True
    app.ScreenUpdating = True

    tgt_name = active_workbook.Name
    active_workbook.Sheets("Contents").Select()

    # Prompt for source report file and open it as read-only.
    src_name = mygetfile(app)
    if str(src_name).lower() == "false":
        return

    app.StatusBar = f"Opening {src_name}"
    app.ScreenUpdating = False
    workbooks.Open(src_name, 0, True)

    for wb in workbooks:
        tst = wb.Name
        if not in_array_helper(open_wbs, tst):
            srcWB = wb
        if tst == tgt_name:
            tgtWB = wb

    # Ensure both source and target workbooks are resolved.
    if srcWB is None or tgtWB is None:
        msg_box("Unable to resolve source or target workbook.", VB_OK_ONLY, "")
        return

    srcWB.Activate()
    bad_report = False

    src_size = str(srcWB.Sheets("Contents").Range("B39")).upper()
    tgt_size = str(tgtWB.Sheets("Contents").Range("B39")).upper()
    if str(srcWB.Sheets("Contents").Range("B50")).startswith("Version"):
        src_version = 2
    else:
        src_version = 3

    if (tgt_size == "SMALL" and tgt_size != src_size) or (tgt_size == "MEDIUM" and src_size == "LARGE"):
        msg = (
            "You are about to import from a larger sized Report Form. Not all data may be imported and the Report may no longer be in balance.  Do you wish to continue?"
        )
        if msg_box(msg, style, "Continue") != VB_OK:
            bad_report = True

    app.StatusBar = f"Validating {srcWB.Name} will fit in this report..."
    if src_version == 3:
        if not bad_report and srcWB.Sheets("Contents").Range("C15") != tgtWB.Sheets("Contents").Range("C15"):
            bad_report = True
            msg_box("Corporate/Subsidiary status does not match. Ending...", VB_OK_ONLY, "")

    if bad_report:
        return

    app.Calculation = XL_MANUAL

    if src_version == 3:
        # Version 3 supports same-sheet-name imports via a generic page copy.
        tgt_sheets = [tgtWB.Sheets(i + 1).Name for i in range(tgtWB.Sheets.Count)]
        src_sheets = [srcWB.Sheets(i + 1).Name for i in range(srcWB.Sheets.Count)]
        for rpt_page in tgt_sheets:
            importPage(rpt_page, not in_array_helper(src_sheets, rpt_page))
    else:
        # Version 2 requires explicit mapping logic into Version 3 layouts.
        importOldtoNew("Contents", "Contents", 8, 13, "C", 1, -3, False)
        tgtWB.Sheets("Contents").Range("C14") = "USD $"
        importOldtoNew("CONTACT_INFO_1", "CONTACT INFO 4", 10, 35, "D", 5, -1, False)
        doPrimary()
        if src_size != "SMALL":
            importOldtoNew("SECONDARY_ACCOUNTS_2b", "SECONDARY ACCOUNTS 3b", 13, 41, "D", 4, -1, False)
        app.StatusBar = "BALANCE_3..."
        tgtWB.Sheets("BALANCE_3").Range("G19") = 0
        tgtWB.Sheets("BALANCE_3").Range("G20") = 0
        if tgtWB.Sheets("BALANCE_3").Range("G31").Locked is False:
            tgtWB.Sheets("BALANCE_3").Range("G31") = 0
        tgtWB.Sheets("BALANCE_3").Range("G19") = srcWB.Sheets("BALANCE 1").Range("G17")
        tgtWB.Sheets("BALANCE_3").Range("G20") = srcWB.Sheets("BALANCE 1").Range("G18")
        if tgtWB.Sheets("BALANCE_3").Range("G31").Locked is False:
            tgtWB.Sheets("BALANCE_3").Range("G31") = srcWB.Sheets("BALANCE 1").Range("G28")
        importOldtoNew("INCOME_4", "INCOME 2", 29, 41, "G", 3, -1, False)
        tgtWB.Sheets("INCOME_4").Range("J18") = 0
        tgtWB.Sheets("INCOME_4").Range("J18") = srcWB.Sheets("INCOME 2").Range("J17")
        doAssets()
        doLiabs()
        importOldtoNew("TRANSFER_IN_9", "TRANSFER IN 9", 13, 57, "C", 4, -4, False)
        importOldtoNew("TRANSFER_OUT_10", "TRANSFER OUT 10", 11, 50, "C", 4, -1, False)
        importOldtoNew("INCOME_DTL_11a", "INCOME DTL 11a", 11, 51, "C", 4, -1, False)
        doIncomeB()
        importOldtoNew("EXPENSE_DTL_12a", "EXPENSE DTL 12a", 12, 54, "C", 4, -1, False)
        doExpenseB()
        importOldtoNew("FINANCE_COMM_13", "FINANCE COMM 13", 11, 54, "C", 4, -1, False)
        importOldtoNew("COMMENTS", "COMMENTS", 8, 32, "C", 1, -1, False)

    # Copy the Free Form page unless running in OpenOffice.
    if not isOO:
        doFreeForm()
    else:
        msg_box("It appears you are runnning Open Office. You will have to transfer any data for the Free Form page manually!", VB_OK_ONLY, "")

    tgtWB.Activate()
    tgtWB.Sheets("Contents").Select()
    app.StatusBar = f"Closing {srcWB.Name}"
    srcWB.Saved = True
    srcWB.Close()

    # Build a new file name based on branch/year/quarter metadata.
    new_name = tgtWB.Sheets("Contents").Range("C8")
    if new_name == "":
        new_name = "Unnamed Branch"
    new_name = (
        f"IMP_RPT_{sanitize(new_name)}_{tgtWB.Sheets('Contents').Range('C11')}_Q{tgtWB.Sheets('Contents').Range('C12')}"
    )
    app.StatusBar = f"Saving {new_name}"
    from . import Module4

    Module4.mysavefile(app, tgtWB, new_name)

    app.Calculation = XL_AUTOMATIC
    tgtWB.Calculate()
    app.DisplayStatusBar = False
    app.ScreenUpdating = True
    msg_box("Done!", VB_OK_ONLY, "")


def importPage(rptPage: str, BlankIt: bool) -> None:
    if rptPage == "Contents":
        start_row, end_row, col, num_cols = 8, 14, "C", 1
    elif rptPage == "CONTACT_INFO_1":
        start_row, end_row, col, num_cols = 10, 35, "D", 5
    elif rptPage == "PRIMARY_ACCOUNT_2a":
        start_row, end_row, col, num_cols = 13, 51, "C", 7
    elif rptPage == "SECONDARY_ACCOUNTS_2b":
        start_row, end_row, col, num_cols = 13, 41, "D", 4
    elif rptPage == "BALANCE_3":
        start_row, end_row, col, num_cols = 19, 31, "G", 1
    elif rptPage == "INCOME_4":
        start_row, end_row, col, num_cols = 18, 41, "G", 4
    elif rptPage == "ASSET_DTL_5a":
        start_row, end_row, col, num_cols = 15, 59, "C", 5
    elif rptPage == "LIABILITY_DTL_5b":
        start_row, end_row, col, num_cols = 16, 55, "C", 5
    elif rptPage in {"INVENTORY_DTL_6", "INVENTORY_DTL_6b", "INVENTORY_DTL_6c"}:
        start_row, end_row, col, num_cols = 13, 30, "E", 8
    elif rptPage == "REGALIA_SALES_DTL_7":
        start_row, end_row, col, num_cols = 20, 51, "C", 7
    elif rptPage == "DEPR_DTL_8":
        start_row, end_row, col, num_cols = 14, 41, "A", 10
    elif rptPage == "TRANSFER_IN_9":
        start_row, end_row, col, num_cols = 14, 57, "C", 4
    elif rptPage == "TRANSFER_OUT_10":
        start_row, end_row, col, num_cols = 11, 50, "C", 4
    elif rptPage == "INCOME_DTL_11a":
        start_row, end_row, col, num_cols = 11, 51, "C", 3
    elif rptPage == "INCOME_DTL_11b":
        start_row, end_row, col, num_cols = 12, 56, "C", 4
    elif rptPage == "EXPENSE_DTL_12a":
        start_row, end_row, col, num_cols = 12, 54, "C", 4
    elif rptPage == "EXPENSE_DTL_12b":
        start_row, end_row, col, num_cols = 12, 55, "C", 4
    elif rptPage == "FINANCE_COMM_13":
        start_row, end_row, col, num_cols = 11, 53, "C", 4
    elif rptPage == "FUNDS_14":
        start_row, end_row, col, num_cols = 14, 55, "D", 3
    elif rptPage == "NEWSLETTER_15":
        start_row, end_row, col, num_cols = 11, 58, "D", 6
    elif rptPage == "COMMENTS":
        start_row, end_row, col, num_cols = 8, 32, "C", 1
    elif rptPage in {"SECONDARY_ACCOUNTS_2b", "SECONDARY ACCOUNTS 2c"}:
        start_row, end_row, col, num_cols = 13, 41, "D", 4
    elif rptPage == "ASSET_DTL_5c":
        start_row, end_row, col, num_cols = 13, 57, "C", 5
    elif rptPage == "LIABILITY_DTL_5d":
        start_row, end_row, col, num_cols = 11, 55, "C", 5
    elif rptPage == "REGALIA_SALES_DTL_7b":
        start_row, end_row, col, num_cols = 20, 51, "C", 7
    elif rptPage in {"DEPR_DTL_8b", "DEPR_DTL_8c"}:
        start_row, end_row, col, num_cols = 14, 53, "A", 10
    elif rptPage in {
        "TRANSFER_IN_9b",
        "TRANSFER_IN_9c",
        "TRANSFER_OUT_10b",
        "TRANSFER_OUT_10c",
        "TRANSFER_OUT_10d",
    }:
        start_row, end_row, col, num_cols = 11, 53, "C", 4
    elif rptPage == "TRANSFER_IN_9d":
        start_row, end_row, col, num_cols = 11, 54, "C", 4
    else:
        return

    if rptPage != "FreeForm":
        importOldtoNew(rptPage, rptPage, start_row, end_row, col, num_cols, 0, BlankIt)


def doFreeForm() -> None:
    if tgtWB is None or srcWB is None:
        return
    tgtWB.Activate()
    tgtWB.Sheets("FreeForm").Select()
    tgtWB.Sheets("FreeForm").Cells.Select()
    tgtWB.Sheets("FreeForm").Cells.ClearContents()
    tgtWB.Sheets("FreeForm").Cells.Delete(Shift=-4162)
    try:
        srcWB.Activate()
        srcWB.Sheets("FreeForm").Select()
        srcWB.Sheets("FreeForm").Shapes.SelectAll()
        srcWB.Sheets("FreeForm").Shapes.Copy()

        tgtWB.Activate()
        tgtWB.Sheets("FreeForm").Range("A1").Select()
        tgtWB.Sheets("FreeForm").Paste()
    finally:
        tgtWB.Sheets("Contents").Select()
        srcWB.Activate()


def doExpenseB() -> None:
    importOldtoNew("EXPENSE_DTL_12b", "EXPENSE DTL 12b", 12, 21, "D", 3, -1, False)
    importOldtoNew("EXPENSE_DTL_12b", "EXPENSE DTL 12b", 27, 41, "C", 4, -2, False)
    importOldtoNew("EXPENSE_DTL_12b", "EXPENSE DTL 12b", 47, 55, "C", 4, -3, False)

    if srcWB.Sheets("EXPENSE DTL 12b").Range("C53") != "":
        old_txt = tgtWB.Sheets("EXPENSE_DTL_12b").Range("C55")
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("C55") = f"{old_txt}, {srcWB.Sheets('EXPENSE DTL 12b').Range('D53')}"

    if srcWB.Sheets("EXPENSE DTL 12b").Range("E53") != "":
        old_txt = tgtWB.Sheets("EXPENSE_DTL_12b").Range("E55")
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("E55") = f"{old_txt}, {srcWB.Sheets('EXPENSE DTL 12b').Range('E53')}"

    if srcWB.Sheets("EXPENSE DTL 12b").Range("F53") != "":
        old_txt = tgtWB.Sheets("EXPENSE_DTL_12b").Range("F55")
        tgtWB.Sheets("EXPENSE_DTL_12b").Range("F55") = old_txt + srcWB.Sheets("EXPENSE DTL 12b").Range("F53")


def doIncomeB() -> None:
    importOldtoNew("INCOME_DTL_11b", "INCOME DTL 11b", 12, 26, "C", 4, -2, False)

    for row in range(25, 32):
        if srcWB.Sheets("INCOME DTL 11b").Range(f"C{row}") != "":
            old_txt = tgtWB.Sheets("INCOME_DTL_11b").Range("C26")
            tgtWB.Sheets("INCOME_DTL_11b").Range("C26") = f"{old_txt}, {srcWB.Sheets('INCOME DTL 11b').Range(f'C{row}')}"

        if srcWB.Sheets("INCOME DTL 11b").Range(f"D{row}") != "":
            old_txt = tgtWB.Sheets("INCOME_DTL_11b").Range("D26")
            tgtWB.Sheets("INCOME_DTL_11b").Range("D26") = old_txt + srcWB.Sheets("INCOME DTL 11b").Range(f"D{row}")

        if srcWB.Sheets("INCOME DTL 11b").Range(f"E{row}") != "":
            old_txt = tgtWB.Sheets("INCOME_DTL_11b").Range("E6")
            tgtWB.Sheets("INCOME_DTL_11b").Range("E26") = old_txt + srcWB.Sheets("INCOME DTL 11b").Range(f"E{row}")

    for row in range(29, 36):
        for j in range(3):
            col = chr(ord("C") + j)
            tgtWB.Sheets("INCOME_DTL_11b").Range(f"{col}{row}") = ""
            if tgtWB.Sheets("INCOME_DTL_11b").Range(f"{col}{row}") != "":
                tgtWB.Sheets("INCOME_DTL_11b").Range(f"{col}{row}") = 0

    importOldtoNew("INCOME_DTL_11b", "INCOME DTL 11b", 40, 46, "C", 3, -5, False)
    importOldtoNew("INCOME_DTL_11b", "INCOME DTL 11b", 50, 56, "C", 4, -5, False)


def doLiabs() -> None:
    for row in range(16, 31):
        for j in range(4):
            col = chr(ord("C") + j)
            tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") = ""
            if tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") != "":
                tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") = 0

    for row in range(37, 44):
        for j in range(4):
            col = chr(ord("C") + j)
            if tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}").Locked is False and row < 41:
                syncCells("LIABILITY_DTL_5b", "COMP BAL DTL 5", f"{col}{row}", f"{col}{row + 7}")
            else:
                tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") = ""
                if tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") != "":
                    tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") = 0

    for row in range(49, 56):
        for j in range(4):
            col = chr(ord("C") + j)
            if tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}").Locked is False and row < 52:
                syncCells("LIABILITY_DTL_5b", "COMP BAL DTL 5", f"{col}{row}", f"{col}{row + 2}")
            else:
                tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") = ""
                if tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") != "":
                    tgtWB.Sheets("LIABILITY_DTL_5b").Range(f"{col}{row}") = 0


def doLiabsOver(BlankIt: bool) -> None:
    importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 11, 28, "C", 5, -3, True)
    importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 33, 37, "C", 5, 9, BlankIt)
    importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 38, 46, "C", 5, -3, True)
    importOldtoNew("LIABILITY_DTL_5d", "COMP BAL DTL 5b", 51, 55, "C", 5, -1, BlankIt)


def doAssets() -> None:
    importOldtoNew("ASSET_DTL_5a", "COMP BAL DTL 5", 15, 18, "C", 5, -2, False)
    importOldtoNew("ASSET_DTL_5a", "COMP BAL DTL 5", 24, 34, "C", 5, -4, False)

    if srcWB.Sheets("COMP BAL DTL 5").Range("C31") != "":
        old_txt = tgtWB.Sheets("ASSET_DTL_5a").Range("C34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("C34") = f"{old_txt}, {srcWB.Sheets('COMP BAL DTL 5').Range('C31')}"

        old_txt = tgtWB.Sheets("ASSET_DTL_5a").Range("D34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("D34") = f"{old_txt}, {srcWB.Sheets('COMP BAL DTL 5').Range('D31')}"

        old_txt = tgtWB.Sheets("ASSET_DTL_5a").Range("F34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("F34") = old_txt + srcWB.Sheets("COMP BAL DTL 5").Range("F31")

        old_txt = tgtWB.Sheets("ASSET_DTL_5a").Range("G34")
        tgtWB.Sheets("ASSET_DTL_5a").Range("G34") = old_txt + srcWB.Sheets("COMP BAL DTL 5").Range("G31")

    for row in range(41, 46):
        for j in range(5):
            col = chr(ord("C") + j)
            tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}") = ""
            if tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}") != "":
                tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}") = 0

    for row in range(52, 60):
        for j in range(5):
            col = chr(ord("C") + j)
            if tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}").Locked is False and row < 56:
                syncCells("ASSET_DTL_5a", "COMP BAL DTL 5", f"{col}{row}", f"{col}{row - 16}")
            elif tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}").Locked is False:
                tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}") = ""
                if tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}") != "":
                    tgtWB.Sheets("ASSET_DTL_5a").Range(f"{col}{row}") = 0


def doAssetsOver(BlankIt: bool) -> None:
    importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 13, 32, "C", 5, -3, BlankIt)
    importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 39, 43, "C", 5, -3, True)
    importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 50, 55, "C", 5, -17, BlankIt)
    importOldtoNew("ASSET_DTL_5c", "COMP BAL DTL 5b", 56, 57, "C", 5, -3, True)


def doPrimary() -> None:
    for row in range(13, 18):
        col = "E"
        s_col = "D"
        if row == 17:
            col = "F"
            s_col = "E"
        if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}").Locked is False:
            syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", f"{col}{row}", f"{s_col}{row - 1}")
        if row in (15, 16):
            col = "H"
            s_col = "G"
            if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}").Locked is False:
                syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", f"{col}{row}", f"{s_col}{row - 1}")

    if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") == 1:
        tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = "Single Signature"
    elif tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") == 2:
        tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h15") = "Dual Signature"

    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h19") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G17")

    for row in range(21, 24):
        for j in range(7):
            col = chr(ord("C") + j)
            if col == "F":
                s_col = "E"
            elif col == "E":
                s_col = "D"
            elif col == "H":
                s_col = "G"
            else:
                s_col = col
            if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}").Locked is False:
                syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", f"{col}{row}", f"{s_col}{row - 2}")

    for row in range(27, 35):
        for j in range(7):
            col = chr(ord("C") + j)
            if col == "E":
                s_col = "D"
            elif col == "F":
                s_col = "E"
            elif col == "H":
                s_col = "G"
            else:
                s_col = col
            if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}").Locked is False and col not in {"D", "G"}:
                syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", f"{col}{row}", f"{s_col}{row - 3}")
            elif tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}").Locked is False:
                tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}") = ""
                if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}") != "":
                    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}") = 0

    for row in range(32, 34):
        if srcWB.Sheets("PRIMARY ACCOUNT 3a").Range(f"C{row}") != "":
            old_txt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("C34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("C34") = f"{old_txt}, {srcWB.Sheets('PRIMARY ACCOUNT 3a').Range(f'C{row}')}"

        if srcWB.Sheets("PRIMARY ACCOUNT 3a").Range(f"E{row}") != "":
            old_txt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F34") = f"{old_txt}, {srcWB.Sheets('PRIMARY ACCOUNT 3a').Range(f'E{row}')}"

        if srcWB.Sheets("PRIMARY ACCOUNT 3a").Range(f"D{row}") != "":
            old_txt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("E34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("E34") = old_txt + srcWB.Sheets("PRIMARY ACCOUNT 3a").Range(f"D{row}")

        if srcWB.Sheets("PRIMARY ACCOUNT 3a").Range(f"G{row}") != "":
            old_txt = tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("I34")
            tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("I34") = old_txt + srcWB.Sheets("PRIMARY ACCOUNT 3a").Range(f"G{row}")

    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h37") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G36")
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F38") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("E37")
    tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("h40") = srcWB.Sheets("PRIMARY ACCOUNT 3a").Range("G39")
    if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F38") != "Yes":
        tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range("F38") = "No"

    for row in range(44, 52):
        for j in range(7):
            col = chr(ord("C") + j)
            if col == "F":
                s_col = "E"
            elif col == "E":
                s_col = "D"
            elif col == "H":
                s_col = "G"
            else:
                s_col = col
            if tgtWB.Sheets("PRIMARY_ACCOUNT_2a").Range(f"{col}{row}").Locked is False and col not in {"D", "H"}:
                syncCells("PRIMARY_ACCOUNT_2a", "PRIMARY ACCOUNT 3a", f"{col}{row}", f"{s_col}{row - 1}")


def importOldtoNew(
    tPage: str,
    sPage: str,
    firstRow: int,
    lastRow: int,
    startCol: str,
    numCols: int,
    offset: int,
    BlankIt: bool,
) -> None:
    for row in range(firstRow, lastRow + 1):
        for j in range(numCols):
            col = chr(ord(startCol) + j)
            if tgtWB.Sheets(tPage).Range(f"{col}{row}").Locked is False:
                if BlankIt:
                    tgtWB.Sheets(tPage).Range(f"{col}{row}") = ""
                else:
                    syncCells(tPage, sPage, f"{col}{row}", f"{col}{row + offset}")


def syncCells(tPage: str, sPage: str, tCell: str, sCell: str) -> None:
    tgtWB.Sheets(tPage).Range(tCell) = ""
    if tgtWB.Sheets(tPage).Range(tCell) != "":
        tgtWB.Sheets(tPage).Range(tCell) = 0
    tgtWB.Sheets(tPage).Range(tCell) = srcWB.Sheets(sPage).Range(sCell)


def mygetfile(app: Any) -> str:
    try:
        fd = app.FileDialog(3)
    except Exception:
        return "False"

    fd.Title = "Select Ledger Workbook to Import"
    fd.AllowMultiSelect = False
    fd.Filters.Clear()
    fd.Filters.Add("Excel Workbooks (*.xls; *.xlsx; *.xlsm)", "*.xls;*.xlsx;*.xlsm")
    fd.Filters.Add("All Files (*.*)", "*.*")

    if fd.Show() != -1:
        return "False"

    return fd.SelectedItems(1)


def get_file(app: Any) -> str:
    try:
        dlg = app.CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    except Exception:
        return "false"

    dlg.execute()
    try:
        return dlg.Files(0)
    except Exception:
        return "false"


def sanitize(fname: str) -> str:
    sanitized = ""
    last_char = ""
    for char in fname:
        code = ord(char)
        if 96 < code < 123 or 64 < code < 91 or 47 < code < 58:
            sanitized += char
            last_char = char
        elif char in {".", "-", "&", "(", ")", "[", "]"}:
            sanitized += char
            last_char = char
        elif last_char != "_":
            sanitized += "_"
            last_char = "_"
    return sanitized


def inArray(values: Iterable[Any], target: Any) -> bool:
    return in_array_helper(values, target)
