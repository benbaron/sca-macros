"""Translated ledger import routines from Module2.bas."""

from __future__ import annotations

from typing import Any, List

from .constants import (
    VB_DEFAULT_BUTTON1,
    VB_EXCLAMATION,
    VB_OK,
    VB_OK_CANCEL,
    VB_OK_ONLY,
    VB_YES,
    VB_YES_NO,
)
from .helpers import msg_box
from . import Module3

PWORD = "SCoE"
LPWORD = "KCoE"


def importfromledger(app: Any, workbook: Any, workbooks: Any) -> None:
    report_version = workbook.Sheets("Contents").Range("B39")
    report_sub = workbook.Sheets("Contents").Range("C15")

    if report_version == "PAYPAL":
        msg_box("This doesn't apply to PAYPAL form.", VB_OK_ONLY, "")
        return
    if report_sub == "Non-US":
        msg_box("Can't import to Non-US report form.", VB_OK_ONLY, "")
        return

    msg = (
        "You are about to import from a ledger form! This may overwrite ALL UNSAVED data already in this workbook."
    )
    msg += "\n\nThe report will be saved in a new file based on the imported ledger's branch name."
    title = "IMPORT Ledger"
    style = VB_OK_CANCEL + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    if msg_box(msg, style, title) != VB_OK:
        return

    qtrarray = [False, False, False, False, False]
    quarter_value = workbook.Sheets("Contents").Range("C12").Value
    if quarter_value == 1:
        qtrarray[1] = True
        msg = "First quarter report. "
    elif quarter_value == 2:
        qtrarray[2] = True
        msg = "Second quarter report. "
        if workbook.Sheets("Contents").Range("C13").Value == "Cumulative":
            qtrarray[1] = True
    elif quarter_value == 3:
        qtrarray[3] = True
        msg = "Third quarter report. "
        if workbook.Sheets("Contents").Range("C13").Value == "Cumulative":
            qtrarray[1] = True
            qtrarray[2] = True
    else:
        qtrarray[4] = True
        msg = "Fourth quarter report. "
        if workbook.Sheets("Contents").Range("C13").Value == "Cumulative":
            qtrarray[1] = True
            qtrarray[2] = True
            qtrarray[3] = True

    msg += " ONLY Import these Ledger Quarters: "
    if qtrarray[1]:
        msg += "First "
    if qtrarray[2]:
        msg += "Second "
    if qtrarray[3]:
        msg += "Third "
    if qtrarray[4]:
        msg += "Fourth "
    if msg_box(msg, style, title) != VB_OK:
        return

    app.DisplayStatusBar = True
    app.ScreenUpdating = True
    app.DisplayAlerts = False

    reportname = workbook.Name
    workbook.Sheets("Contents").Select()

    ledgername = Module3.mygetfile(app)
    if str(ledgername).lower() == "false":
        return

    app.StatusBar = f"Opening {ledgername}"
    app.ScreenUpdating = False
    workbooks.Open(ledgername)

    split_char = "/" if ledgername.startswith("file://") or ledgername.startswith("/") else "\\"
    str_path = ledgername.split(split_char)
    ledgername = str_path[-1]

    ledger_wb = workbooks(ledgername)
    ledger_wb.Activate()
    badledger = False

    ledger_vers = ledger_wb.Sheets("Contents").Range("F46").Value.split(" ")
    if ledger_vers[-1].startswith("3"):
        lq1 = "Ledger_Q1"
        lq2 = "Ledger_Q2"
        lq3 = "Ledger_Q3"
        lq4 = "Ledger_Q4"
    else:
        ledger_wb.Close()
        workbooks(reportname).Activate()
        msg_box("Please convert your source Ledger to version 3 to enable Import.", VB_OK_ONLY, "")
        return

    ledger_wb.Sheets("Contents").Select()
    ledger_wb.Unprotect(LPWORD)
    ledger_wb.Sheets("Summary").Visible = True
    ledger_wb.Sheets(lq1).Visible = True
    ledger_wb.Sheets(lq2).Visible = True
    ledger_wb.Sheets(lq3).Visible = True
    ledger_wb.Sheets(lq4).Visible = True
    ledger_wb.Protect(LPWORD)

    app.StatusBar = f"Unlocking {ledgername}"
    for sheet in ledger_wb.Worksheets:
        sheet.Unprotect(LPWORD)

    if (
        workbooks(reportname).Sheets("Contents").Range("C8") != ""
        or workbooks(reportname).Sheets("Contents").Range("C11") > 0
    ):
        if (
            workbooks(reportname).Sheets("Contents").Range("C8")
            != workbooks(ledgername).Sheets("Contents").Range("C4")
        ):
            msg_box("Branch name does not match. Ending...", VB_OK_ONLY, "")
            badledger = True
        if not badledger and workbooks(reportname).Sheets("Contents").Range("C11") > 0:
            if (
                workbooks(reportname).Sheets("Contents").Range("C11")
                != workbooks(ledgername).Sheets("Contents").Range("C5")
            ):
                badledger = True
                msg_box("Year does not match. Ending...", VB_OK_ONLY, "")
        if not badledger and ledger_wb.Sheets("Contents").Range("C6") == report_sub and report_sub == "Corporate":
            pass
        elif not badledger and report_sub == ledger_wb.Sheets("Contents").Range("C6"):
            pass
        else:
            badledger = True
            msg_box("Corporate/Subsidiary status does not match. Ending...", VB_OK_ONLY, "")

    app.StatusBar = f"Validating {ledgername} will fit in this report..."
    if not badledger and report_version == "SMALL":
        if workbooks(ledgername).Sheets("Summary").Range("C15").Value != "":
            badledger = True
            msg_box("SMALL form does not have enough room for the secondary accounts, use MEDIUM form...", VB_OK_ONLY, "")
        if not badledger:
            summary = workbooks(ledgername).Sheets("Summary")
            if not (
                summary.Range("D27") == 0
                and summary.Range("E27") == 0
                and summary.Range("D28") == 0
                and summary.Range("E28") == 0
                and summary.Range("D29") == 0
                and summary.Range("E29") == 0
            ):
                badledger = True
                msg_box("SMALL form does not have room for assets, use MEDIUM form...", VB_OK_ONLY, "")
        if not badledger:
            summary = workbooks(ledgername).Sheets("Summary")
            if not (summary.Range("D32") == 0 and summary.Range("E32") == 0):
                badledger = True
                msg_box("SMALL form does not have room for newsletters, use MEDIUM form...", VB_OK_ONLY, "")

    if not badledger:
        app.Calculation = -4135
        import_ledger_pages(app, workbook, workbooks, ledgername, qtrarray)
        app.Calculation = -4105
        app.DisplayStatusBar = False
        app.ScreenUpdating = True
        msg_box("Done!", VB_OK_ONLY, "")


def import_ledger_pages(app: Any, workbook: Any, workbooks: Any, ledgername: str, qtrarray: List[bool]) -> None:
    ledger_wb = workbooks(ledgername)
    app.StatusBar = "Importing data..."

    ledger_report_map = [
        ("BALANCE_3", "Summary", "G19", "D8"),
        ("BALANCE_3", "Summary", "G20", "E8"),
    ]
    for t_sheet, s_sheet, t_cell, s_cell in ledger_report_map:
        workbook.Sheets(t_sheet).Range(t_cell).Value = ledger_wb.Sheets(s_sheet).Range(s_cell).Value

    if workbook.Sheets("Contents").Range("B39") != "SMALL":
        workbook.Sheets("BALANCE_3").Range("G31").Value = ledger_wb.Sheets("Summary").Range("E14").Value

    _import_qtr(app, workbook, ledger_wb, "Ledger_Q1", qtrarray[1])
    _import_qtr(app, workbook, ledger_wb, "Ledger_Q2", qtrarray[2])
    _import_qtr(app, workbook, ledger_wb, "Ledger_Q3", qtrarray[3])
    _import_qtr(app, workbook, ledger_wb, "Ledger_Q4", qtrarray[4])

    app.StatusBar = "Syncing report totals..."
    workbook.Calculate()


def _import_qtr(app: Any, workbook: Any, ledger_wb: Any, sheet_name: str, enabled: bool) -> None:
    if not enabled:
        return
    app.StatusBar = f"Importing {sheet_name}..."
    src_sheet = ledger_wb.Sheets(sheet_name)

    workbook.Sheets("INCOME_4").Range("G19").Value = src_sheet.Range("L17").Value
    workbook.Sheets("INCOME_4").Range("G20").Value = src_sheet.Range("L18").Value
    workbook.Sheets("INCOME_4").Range("G21").Value = src_sheet.Range("L19").Value
    workbook.Sheets("INCOME_4").Range("G22").Value = src_sheet.Range("L20").Value
    workbook.Sheets("INCOME_4").Range("G23").Value = src_sheet.Range("L21").Value
    workbook.Sheets("INCOME_4").Range("G24").Value = src_sheet.Range("L22").Value
    workbook.Sheets("INCOME_4").Range("G25").Value = src_sheet.Range("L23").Value
    workbook.Sheets("INCOME_4").Range("G26").Value = src_sheet.Range("L24").Value
    workbook.Sheets("INCOME_4").Range("G27").Value = src_sheet.Range("L25").Value

    workbook.Sheets("INCOME_4").Range("H19").Value = src_sheet.Range("M17").Value
    workbook.Sheets("INCOME_4").Range("H20").Value = src_sheet.Range("M18").Value
    workbook.Sheets("INCOME_4").Range("H21").Value = src_sheet.Range("M19").Value
    workbook.Sheets("INCOME_4").Range("H22").Value = src_sheet.Range("M20").Value
    workbook.Sheets("INCOME_4").Range("H23").Value = src_sheet.Range("M21").Value
    workbook.Sheets("INCOME_4").Range("H24").Value = src_sheet.Range("M22").Value
    workbook.Sheets("INCOME_4").Range("H25").Value = src_sheet.Range("M23").Value
    workbook.Sheets("INCOME_4").Range("H26").Value = src_sheet.Range("M24").Value
    workbook.Sheets("INCOME_4").Range("H27").Value = src_sheet.Range("M25").Value
