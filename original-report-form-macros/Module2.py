"""Ledger import routines translated from the original VBA Module2.

This module handles importing ledger data into a report workbook, validating
ledger compatibility, and mapping ledger values into report sheets.
"""

from __future__ import annotations

from typing import Any, List, Sequence

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
from . import Module3, Module4, Module6

PWORD = "SCoE"
LPWORD = "KCoE"


def importfromledger(app: Any, workbook: Any, workbooks: Any) -> None:
    """Import ledger data into the active report workbook."""
    report_version = workbook.Sheets("Contents").Range("B39")
    report_sub = workbook.Sheets("Contents").Range("C15")

    # Guardrails: PayPal and Non-US forms are not compatible with imports.
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

    # Determine which quarters can be imported based on report quarter/cumulative mode.
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

    # Prompt for ledger file path, then open the source ledger workbook.
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

    # Verify ledger version for compatibility with the report importer.
    ledger_vers = ledger_wb.Sheets("Contents").Range("F46").Value.split(" ")
    if ledger_vers[-1].startswith("3"):
        lq1 = "Ledger_Q1"
        lq2 = "Ledger_Q2"
        lq3 = "Ledger_Q3"
        lq4 = "Ledger_Q4"
        eql = "Equipment_List"
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

    # Validate branch name, year, and corporate/subsidiary status.
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

    # Confirm the ledger fits within the target report size before import.
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
            if getledgervalue(ledger_wb, "AX19", qtrarray) > 24 or getledgervalue(ledger_wb, "AX20", qtrarray) > 16:
                badledger = True
                msg_box("SMALL form does not have enough room for the transfers in, use MEDIUM form...", VB_OK_ONLY, "")

    if not badledger:
        app.Calculation = -4135
        import_ledger_pages(
            app,
            workbook,
            workbooks,
            ledgername,
            qtrarray,
            report_version,
            report_sub,
            eql,
            title,
        )
        app.Calculation = -4105
        app.DisplayStatusBar = False
        app.ScreenUpdating = True
        app.DisplayAlerts = True
        msg_box("Done!", VB_OK_ONLY, "")
        return

    app.StatusBar = f"Closing {ledgername}"
    ledger_wb.Saved = True
    ledger_wb.Close()
    workbooks(reportname).Activate()
    app.DisplayStatusBar = False
    app.ScreenUpdating = True
    app.DisplayAlerts = True


def import_ledger_pages(
    app: Any,
    workbook: Any,
    workbooks: Any,
    ledgername: str,
    qtrarray: List[bool],
    report_version: str,
    report_sub: str,
    eql: str,
    title: str,
) -> None:
    """Copy ledger values into the report workbook for selected quarters."""
    ledger_wb = workbooks(ledgername)
    app.StatusBar = "Importing data..."

    # 4. Fill in contents.
    if workbook.Sheets("Contents").Range("C8") == "":
        lyear = ledger_wb.Sheets("Contents").Range("C5")
        workbook.Sheets("Contents").Range("C8").Value = ledger_wb.Sheets("Contents").Range("C4")
        workbook.Sheets("Contents").Range("C10").Value = ledger_wb.Sheets("Signatories").Range("D5")
        workbook.Sheets("Contents").Range("C11").Value = ledger_wb.Sheets("Contents").Range("C5")
        if ledger_wb.Sheets("Contents").Range("C6") != "Corporate":
            workbook.Sheets("Contents").Range("C15").Value = ledger_wb.Sheets("Contents").Range("C6")
    lyear = workbook.Sheets("Contents").Range("C11")

    # 4a. Exchequer information from signatory page.
    app.StatusBar = "Contact Info..."
    workbook.Sheets("CONTACT_INFO_1").Range("D12").Value = ledger_wb.Sheets("Signatories").Range("E5")
    if ledger_wb.Sheets("Signatories").Range("E6") != "":
        addrstr = ledger_wb.Sheets("Signatories").Range("E6")
        commaloc = str(addrstr).find(",")
        if commaloc > -1:
            workbook.Sheets("CONTACT_INFO_1").Range("D13").Value = str(addrstr)[: commaloc + 1]
            addrstr = str(addrstr)[commaloc + 1 :]
            addrstr = str(addrstr).lstrip()
            csz = addrstr.split(" ")
            workbook.Sheets("CONTACT_INFO_1").Range("H13").Value = csz[-1]
            if len(csz) == 2:
                workbook.Sheets("CONTACT_INFO_1").Range("F13").Value = csz[0]
            else:
                workbook.Sheets("CONTACT_INFO_1").Range("F13").Value = addrstr[: addrstr.find(csz[-1])]
        else:
            workbook.Sheets("CONTACT_INFO_1").Range("D13").Value = addrstr
    workbook.Sheets("CONTACT_INFO_1").Range("D14").Value = ledger_wb.Sheets("Signatories").Range("E8")
    workbook.Sheets("CONTACT_INFO_1").Range("F14").Value = ledger_wb.Sheets("Signatories").Range("F8")
    workbook.Sheets("CONTACT_INFO_1").Range("D15").Value = ledger_wb.Sheets("Signatories").Range("D8")
    workbook.Sheets("CONTACT_INFO_1").Range("D16").Value = ledger_wb.Sheets("Signatories").Range("D6")
    workbook.Sheets("CONTACT_INFO_1").Range("H15").Value = ledger_wb.Sheets("Signatories").Range("F5")
    workbook.Sheets("CONTACT_INFO_1").Range("H16").Value = ledger_wb.Sheets("Signatories").Range("F6")

    # 5. Fill in accounts.
    app.StatusBar = "Accounts..."
    if qtrarray[4]:
        workbook.Sheets("Primary_Account_2a").Range("h37").Value = ledger_wb.Sheets("Ledger_Q4").Range("c110").Value
        workbook.Sheets("Primary_Account_2a").Range("h19").Value = ledger_wb.Sheets("Balances").Range("N5").Value
    elif qtrarray[3]:
        workbook.Sheets("Primary_Account_2a").Range("h37").Value = ledger_wb.Sheets("Ledger_Q3").Range("c110").Value
        workbook.Sheets("Primary_Account_2a").Range("h19").Value = ledger_wb.Sheets("Balances").Range("K5").Value
    elif qtrarray[2]:
        workbook.Sheets("Primary_Account_2a").Range("h37").Value = ledger_wb.Sheets("Ledger_Q2").Range("c110").Value
        workbook.Sheets("Primary_Account_2a").Range("h19").Value = ledger_wb.Sheets("Balances").Range("H5").Value
    elif qtrarray[1]:
        workbook.Sheets("Primary_Account_2a").Range("h37").Value = ledger_wb.Sheets("Ledger_Q1").Range("c110").Value
        workbook.Sheets("Primary_Account_2a").Range("h19").Value = ledger_wb.Sheets("Balances").Range("E5").Value

    workbook.Sheets("Primary_Account_2a").Range("e15").Value = ledger_wb.Sheets("Balances").Range("b3").Value
    workbook.Sheets("Primary_Account_2a").Range("e16").Value = ledger_wb.Sheets("Balances").Range("b4").Value
    workbook.Sheets("Primary_Account_2a").Range("h15").Value = ledger_wb.Sheets("Balances").Range("b5").Value
    workbook.Sheets("Primary_Account_2a").Range("f38").Value = ledger_wb.Sheets("Balances").Range("b7").Value
    workbook.Sheets("Primary_Account_2a").Range("e13").Value = ledger_wb.Sheets("Balances").Range("b8").Value
    workbook.Sheets("Primary_Account_2a").Range("e14").Value = ledger_wb.Sheets("Balances").Range("e8").Value
    workbook.Sheets("Primary_Account_2a").Range("f17").Value = ledger_wb.Sheets("Balances").Range("c9").Value

    # Clear existing outstanding checks.
    for i in range(27, 35):
        for col in range(3, 9):
            workbook.Sheets("Primary_Account_2a").Cells(i, col).Value = ""

    months_array = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    if qtrarray[4]:
        pass
    elif qtrarray[3]:
        months_array = months_array[:9]
    elif qtrarray[2]:
        months_array = months_array[:6]
    else:
        months_array = months_array[:3]

    lrow = 4
    lcol = 16
    rrow = 27
    rcol = 3

    for _ in range(1, 31):
        if not Module3.inArray(months_array, ledger_wb.Sheets("Balances").Cells(lrow, lcol + 3).Value):
            if ledger_wb.Sheets("Balances").Cells(lrow, lcol + 2).Value < 0:
                if workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol + 2).Value == "":
                    workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol).Value = ledger_wb.Sheets("Balances").Cells(
                        lrow, lcol
                    ).Value
                    workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol + 1).Value = ledger_wb.Sheets(
                        "Balances"
                    ).Cells(lrow, lcol + 1).Value
                    workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol + 2).Value = (
                        ledger_wb.Sheets("Balances").Cells(lrow, lcol + 2).Value * -1
                    )
                    rrow += 1
                else:
                    tst = str(workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol).Value).split(" ")
                    workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol).Value = (
                        f"{tst[0]} - {ledger_wb.Sheets('Balances').Cells(lrow, lcol).Value}"
                    )
                    workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol + 2).Value = (
                        workbook.Sheets("Primary_Account_2a").Cells(rrow, rcol + 2).Value
                        + (ledger_wb.Sheets("Balances").Cells(lrow, lcol + 2).Value * -1)
                    )

                if rrow == 35:
                    if rcol == 6:
                        rrow = 34
                    else:
                        rrow = 27
                        rcol += 3
        lrow += 1
        if lrow == 10:
            lrow = 4
            lcol += 5

    if rrow < 35:
        getoutstandingchecks(ledger_wb, workbook, rrow, rcol, qtrarray, months_array)

    rsrow = 44
    rscol = 3
    lsrow = 9
    lscol = 3
    for _ in range(2, 21):
        if ledger_wb.Sheets("Signatories").Cells(lsrow, lscol + 4).Value == "X":
            workbook.Sheets("Primary_Account_2a").Cells(rsrow, rscol).Value = ledger_wb.Sheets("Signatories").Cells(
                lsrow, lscol
            ).Value
            workbook.Sheets("Primary_Account_2a").Cells(rsrow, rscol + 2).Value = ledger_wb.Sheets(
                "Signatories"
            ).Cells(lsrow, lscol + 1).Value
            workbook.Sheets("Primary_Account_2a").Cells(rsrow, rscol + 3).Value = ledger_wb.Sheets(
                "Signatories"
            ).Cells(lsrow, lscol + 2).Value
            workbook.Sheets("Primary_Account_2a").Cells(rsrow + 1, rscol + 3).Value = ledger_wb.Sheets(
                "Signatories"
            ).Cells(lsrow + 1, lscol + 2).Value
            workbook.Sheets("Primary_Account_2a").Cells(rsrow, rscol + 5).Value = ledger_wb.Sheets(
                "Signatories"
            ).Cells(lsrow, lscol + 3).Value
            workbook.Sheets("Primary_Account_2a").Cells(rsrow + 1, rscol + 5).Value = ledger_wb.Sheets(
                "Signatories"
            ).Cells(lsrow + 1, lscol + 3).Value
            rsrow += 2
            lsrow += 4
        if rsrow == 54:
            break

    # Secondary accounts.
    workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range("D13:g21,D25:g25,D27:g41").ClearContents()
    if report_version in {"LARGE", "MASTER"}:
        workbook.Sheets("SECONDARY_ACCOUNTS_2c").Range("D13:g21,D25:g25,D27:g41").ClearContents()
        workbook.Sheets("SECONDARY_ACCOUNTS_2d").Range("D13:g21,D25:g25,D27:g41").ClearContents()

    lrow = 14
    lcol = 16
    rrow = 13
    rcol = 4
    tsheetname = "Secondary_Accounts_2b"
    for i in range(11, 23):
        if ledger_wb.Sheets("Summary").Cells(i, 3).Value == "":
            break
        if report_version != "LARGE" and i == 15:
            break
        workbook.Sheets(tsheetname).Cells(13, rcol).Value = ledger_wb.Sheets("Summary").Cells(i, 3).Value
        workbook.Sheets(tsheetname).Cells(16, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow - 1, lcol - 14)
        workbook.Sheets(tsheetname).Cells(14, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow, lcol - 14)
        workbook.Sheets(tsheetname).Cells(15, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow + 1, lcol - 14)
        workbook.Sheets(tsheetname).Cells(17, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow + 3, lcol - 14)
        workbook.Sheets(tsheetname).Cells(21, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow, 42)
        if qtrarray[4]:
            workbook.Sheets(tsheetname).Cells(19, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow + 1, lcol - 2)
            workbook.Sheets(tsheetname).Cells(25, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow - 1, lcol - 2)
        elif qtrarray[3]:
            workbook.Sheets(tsheetname).Cells(19, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow + 1, lcol - 5)
            workbook.Sheets(tsheetname).Cells(25, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow - 1, lcol - 5)
        elif qtrarray[2]:
            workbook.Sheets(tsheetname).Cells(19, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow + 1, lcol - 8)
            workbook.Sheets(tsheetname).Cells(25, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow - 1, lcol - 8)
        else:
            workbook.Sheets(tsheetname).Cells(19, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow + 1, lcol - 11)
            workbook.Sheets(tsheetname).Cells(25, rcol).Value = ledger_wb.Sheets("Balances").Cells(lrow - 1, lcol - 11)

        rsrow = 27
        lscol = 4
        lsrow = 5
        for _ in range(1, 21):
            if ledger_wb.Sheets("Signatories").Cells(lsrow, i - 3).Value == "X":
                workbook.Sheets(tsheetname).Cells(rsrow, rcol).Value = ledger_wb.Sheets("Signatories").Cells(
                    lsrow, lscol
                ).Value
                workbook.Sheets(tsheetname).Cells(rsrow + 1, rcol).Value = ledger_wb.Sheets("Signatories").Cells(
                    lsrow, lscol + 2
                ).Value
                workbook.Sheets(tsheetname).Cells(rsrow + 2, rcol).Value = ledger_wb.Sheets("Signatories").Cells(
                    lsrow + 1, lscol + 2
                ).Value
                rsrow += 3
            lsrow += 4
            if rsrow == 45:
                break

        rcol += 1
        lrow += 10

        if i == 14:
            tsheetname = "Secondary_Accounts_2c"
            rcol = 4
        elif i == 18:
            tsheetname = "Secondary_Accounts_2d"
            rcol = 4

    # 6a. Balance statement previous balance.
    noninttot = 0
    inttot = 0
    lrow = 6
    ldgradd = 10
    ldgraddval = 0
    for i in range(1, 14):
        if not qtrarray[1]:
            if not qtrarray[2]:
                if not qtrarray[3]:
                    ldgraddval = (
                        ledger_wb.Sheets("Ledger_Q1").Range(f"AO{i + ldgradd}").Value
                        + ledger_wb.Sheets("Ledger_Q2").Range(f"AO{i + ldgradd}").Value
                        + ledger_wb.Sheets("Ledger_Q3").Range(f"AO{i + ldgradd}").Value
                    )
                else:
                    ldgraddval = (
                        ledger_wb.Sheets("Ledger_Q1").Range(f"AO{i + ldgradd}").Value
                        + ledger_wb.Sheets("Ledger_Q2").Range(f"AO{i + ldgradd}").Value
                    )
            else:
                ldgraddval = ledger_wb.Sheets("Ledger_Q1").Range(f"AO{i + ldgradd}").Value
        if ledger_wb.Sheets("Balances").Cells(lrow + 1, 2).Value == "YES":
            inttot += ledger_wb.Sheets("Summary").Range(f"D{i + 9}").Value + ldgraddval
        elif ledger_wb.Sheets("Balances").Cells(lrow + 1, 2).Value == "NO":
            noninttot += ledger_wb.Sheets("Summary").Range(f"D{i + 9}").Value + ldgraddval
        lrow += 10
        if i == 1:
            ldgradd = 20
    workbook.Sheets("BALANCE_3").Range("G19").Value = noninttot
    workbook.Sheets("BALANCE_3").Range("G20").Value = inttot

    # Summary data mappings that feed the balance sheet.
    ledger_report_map = [
        ("BALANCE_3", "Summary", "G19", "D8"),
        ("BALANCE_3", "Summary", "G20", "E8"),
    ]
    for t_sheet, s_sheet, t_cell, s_cell in ledger_report_map:
        workbook.Sheets(t_sheet).Range(t_cell).Value = ledger_wb.Sheets(s_sheet).Range(s_cell).Value

    if report_version != "SMALL":
        workbook.Sheets("BALANCE_3").Range("G31").Value = ledger_wb.Sheets("Summary").Range("E14").Value

    # 6b. Income statement.
    app.StatusBar = "Income Statement..."
    Module6.ClearIncomeExpense(workbook, report_version)

    workbook.Sheets("INCOME_4").Range("J18").Value = getledgervalue(ledger_wb, "AS21", qtrarray)
    workbook.Sheets("INCOME_DTL_11a").Range("E33").Value = getledgervalue(ledger_wb, "AS14", qtrarray)
    workbook.Sheets("INCOME_DTL_11a").Range("E34").Value = getledgervalue(ledger_wb, "AS15", qtrarray)
    workbook.Sheets("INCOME_DTL_11a").Range("E35").Value = getledgervalue(ledger_wb, "AS16", qtrarray)
    if report_version not in {"SMALL", "PAYPAL"}:
        workbook.Sheets("NEWSLETTER_15").Range("I11").Value = getledgervalue(ledger_wb, "AS24", qtrarray)

    workbook.Sheets("INCOME_4").Range("G29").Value = getledgervalue(ledger_wb, "AU15", qtrarray)
    workbook.Sheets("INCOME_4").Range("H29").Value = getledgervalue(ledger_wb, "AU16", qtrarray)
    workbook.Sheets("INCOME_4").Range("I29").Value = getledgervalue(ledger_wb, "AU17", qtrarray)

    workbook.Sheets("INCOME_4").Range("G31").Value = getledgervalue(ledger_wb, "AU18", qtrarray)
    workbook.Sheets("INCOME_4").Range("H31").Value = getledgervalue(ledger_wb, "AU19", qtrarray)
    workbook.Sheets("INCOME_4").Range("I31").Value = getledgervalue(ledger_wb, "AU20", qtrarray)

    workbook.Sheets("INCOME_4").Range("G33").Value = getledgervalue(ledger_wb, "AU24", qtrarray)
    workbook.Sheets("INCOME_4").Range("H33").Value = getledgervalue(ledger_wb, "AU25", qtrarray)
    workbook.Sheets("INCOME_4").Range("I33").Value = getledgervalue(ledger_wb, "AU26", qtrarray)

    workbook.Sheets("INCOME_4").Range("G34").Value = getledgervalue(ledger_wb, "AU27", qtrarray)
    workbook.Sheets("INCOME_4").Range("H34").Value = getledgervalue(ledger_wb, "AU28", qtrarray)
    workbook.Sheets("INCOME_4").Range("I34").Value = getledgervalue(ledger_wb, "AU29", qtrarray)

    workbook.Sheets("INCOME_4").Range("G36").Value = getledgervalue(ledger_wb, "AU31", qtrarray)
    workbook.Sheets("INCOME_4").Range("H36").Value = getledgervalue(ledger_wb, "AU32", qtrarray)
    workbook.Sheets("INCOME_4").Range("I36").Value = getledgervalue(ledger_wb, "AU33", qtrarray)

    workbook.Sheets("INCOME_4").Range("G37").Value = getledgervalue(ledger_wb, "AU34", qtrarray)
    workbook.Sheets("INCOME_4").Range("H37").Value = getledgervalue(ledger_wb, "AU35", qtrarray)
    workbook.Sheets("INCOME_4").Range("I37").Value = getledgervalue(ledger_wb, "AU36", qtrarray)

    workbook.Sheets("INCOME_4").Range("G38").Value = getledgervalue(ledger_wb, "AU37", qtrarray)
    workbook.Sheets("INCOME_4").Range("H38").Value = getledgervalue(ledger_wb, "AU38", qtrarray)
    workbook.Sheets("INCOME_4").Range("I38").Value = getledgervalue(ledger_wb, "AU39", qtrarray)

    workbook.Sheets("INCOME_4").Range("G40").Value = getledgervalue(ledger_wb, "AU41", qtrarray)
    workbook.Sheets("INCOME_4").Range("H40").Value = getledgervalue(ledger_wb, "AU42", qtrarray)
    workbook.Sheets("INCOME_4").Range("I40").Value = getledgervalue(ledger_wb, "AU43", qtrarray)

    workbook.Sheets("INCOME_4").Range("G41").Value = getledgervalue(ledger_wb, "AU44", qtrarray)
    workbook.Sheets("INCOME_4").Range("H41").Value = getledgervalue(ledger_wb, "AU45", qtrarray)
    workbook.Sheets("INCOME_4").Range("I41").Value = getledgervalue(ledger_wb, "AU46", qtrarray)

    # 7. Assets.
    app.StatusBar = "Assets..."
    workbook.Sheets("ASSET_DTL_5a").Range("c15:g18,c24:g34,c41:g45,c52:g59").ClearContents()
    if report_version in {"LARGE", "PAYPAL", "MASTER"}:
        workbook.Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57").ClearContents()

    if getledgervalue(ledger_wb, "av12", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("An12").Value
        writereportline(ledger_wb, workbook, matchname, "ASSET_DTL_5a", 24, 34, 14, qtrarray)
    if getledgervalue(ledger_wb, "av16", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("An16").Value
        writereportline(ledger_wb, workbook, matchname, "ASSET_DTL_5a", 41, 45, 14, qtrarray)
    if getledgervalue(ledger_wb, "av17", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("An17").Value
        writereportline(ledger_wb, workbook, matchname, "ASSET_DTL_5a", 52, 59, 14, qtrarray)

    # 8. Liabilities.
    app.StatusBar = "Liabilities..."
    workbook.Sheets("LIABILITY_DTL_5b").Range("c16:f30,c37:f43,c49:f55").ClearContents()
    if report_version in {"LARGE", "PAYPAL", "MASTER"}:
        workbook.Sheets("LIABILITY_DTL_5d").Range("c11:f28,c33:f46,c51:f55").ClearContents()
        if report_version in {"PAYPAL", "MASTER"}:
            workbook.Sheets("LIABILITY_DTL_5e").Range("c11:f55").ClearContents()
            workbook.Sheets("LIABILITY_DTL_5f").Range("c11:f55").ClearContents()
            workbook.Sheets("LIABILITY_DTL_5g").Range("c11:f55").ClearContents()

    if getledgervalue(ledger_wb, "av19", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("An19").Value
        writereportline(ledger_wb, workbook, matchname, "LIABILITY_DTL_5b", 16, 30, 14, qtrarray)
    if getledgervalue(ledger_wb, "av20", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("An20").Value
        writereportline(ledger_wb, workbook, matchname, "LIABILITY_DTL_5b", 37, 43, 14, qtrarray)
    if getledgervalue(ledger_wb, "av21", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("An21").Value
        writereportline(ledger_wb, workbook, matchname, "LIABILITY_DTL_5b", 49, 55, 14, qtrarray)

    # Non-cash assets and funds for larger reports.
    if report_version != "SMALL":
        app.StatusBar = "Clearing Non-cash Assets..."
        workbook.Sheets("INVENTORY_DTL_6").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents()
        workbook.Sheets("REGALIA_SALES_DTL_7").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents()
        workbook.Sheets("DEPR_DTL_8").Range("d14:g23,j14:j23,e32:g41,j32:j41").ClearContents()

        if report_version in {"LARGE", "MASTER"}:
            workbook.Sheets("INVENTORY_DTL_6b").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents()
            workbook.Sheets("REGALIA_SALES_DTL_7b").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents()
            workbook.Sheets("DEPR_DTL_8b").Range("d14:g53,j14:j53").ClearContents()
            workbook.Sheets("DEPR_DTL_8c").Range("e14:g53,j14:j53").ClearContents()

        # 9. Depreciation.
        app.StatusBar = "Depreciation..."
        fiveyrtotal = (
            ledger_wb.Sheets(eql).Range("U9").Value
            + ledger_wb.Sheets(eql).Range("U10").Value
            + ledger_wb.Sheets(eql).Range("U11").Value
        )
        if fiveyrtotal > 0:
            msg_box(
                f"There are {fiveyrtotal} 5-year Depreciation Assets marked on the Equipment List.",
                VB_OK_ONLY,
                "",
            )
            if workbook.Sheets("DEPR_DTL_8").Range("E14") != "":
                msg = "Depreciation information is already in this Report Form. Overwrite?"
                style = VB_YES_NO + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
                doitresponse = msg_box(msg, style, title)
            else:
                doitresponse = VB_YES
            if doitresponse == VB_YES:
                reportline = 14
                targetpage = "DEPR_DTL_8"
                for ledgerline in range(11, 261):
                    if "5yr" in str(ledger_wb.Sheets(eql).Cells(ledgerline, 8).Value):
                        workbook.Sheets(targetpage).Cells(reportline, 4).Value = str(
                            ledger_wb.Sheets(eql).Cells(ledgerline, 8).Value
                        )[:2]
                        workbook.Sheets(targetpage).Cells(reportline, 5).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 4
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 6).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 5
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 7).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 6
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 10).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 7
                        ).Value
                        reportline += 1
                        if reportline == 24 and targetpage == "DEPR_DTL_8":
                            if report_version == "LARGE":
                                targetpage = "DEPR_DTL_8b"
                                reportline = 14
                            else:
                                break

        if ledger_wb.Sheets(eql).Range("U12").Value > 0:
            msg_box(
                f"There are {ledger_wb.Sheets(eql).Range('U12').Value} 7-year Depreciation Assets marked on the Equipment List.",
                VB_OK_ONLY,
                "",
            )
            if workbook.Sheets("DEPR_DTL_8").Range("E32") != "":
                msg = "Depreciation information is already in this Report Form. Overwrite?"
                style = VB_YES_NO + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
                doitresponse = msg_box(msg, style, title)
            else:
                doitresponse = VB_YES
            if doitresponse == VB_YES:
                reportline = 32
                targetpage = "DEPR_DTL_8"
                for ledgerline in range(11, 261):
                    if "7yr" in str(ledger_wb.Sheets(eql).Cells(ledgerline, 8).Value):
                        workbook.Sheets(targetpage).Cells(reportline, 5).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 4
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 6).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 5
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 7).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 6
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 10).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 7
                        ).Value
                        reportline += 1
                        if reportline == 41 and targetpage == "DEPR_DTL_8":
                            if report_version == "LARGE":
                                targetpage = "DEPR_DTL_8c"
                                reportline = 14
                            else:
                                break

        # 10. Regalia.
        app.StatusBar = f"Regalia...{ledger_wb.Sheets(eql).Range('U8').Value}"
        if ledger_wb.Sheets(eql).Range("U8").Value > 0:
            msg_box(
                f"There are {ledger_wb.Sheets(eql).Range('U8').Value} Regalia Assets marked on the Equipment List. Only copying those with value >= $500.",
                VB_OK_ONLY,
                "",
            )
            if workbook.Sheets("REGALIA_SALES_DTL_7").Range("C20") != "":
                msg = "Regalia information is already in this Report Form. Overwrite?"
                style = VB_YES_NO + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
                doitresponse = msg_box(msg, style, title)
            else:
                doitresponse = VB_YES
            if doitresponse == VB_YES:
                reportline = 20
                targetpage = "REGALIA_SALES_DTL_7"
                for ledgerline in range(11, 261):
                    if (
                        ledger_wb.Sheets(eql).Cells(ledgerline, 8).Value == "Regalia"
                        and ledger_wb.Sheets(eql).Cells(ledgerline, 7).Value >= 500
                    ):
                        workbook.Sheets(targetpage).Cells(reportline, 3).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 4
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 4).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 5
                        ).Value
                        workbook.Sheets(targetpage).Cells(reportline, 5).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 6
                        ).Value
                        if ledger_wb.Sheets(eql).Cells(ledgerline, 6).Value == lyear:
                            workbook.Sheets(targetpage).Cells(reportline, 7).Value = ledger_wb.Sheets(eql).Cells(
                                ledgerline, 7
                            ).Value
                        else:
                            workbook.Sheets(targetpage).Cells(reportline, 6).Value = ledger_wb.Sheets(eql).Cells(
                                ledgerline, 7
                            ).Value
                        reportline += 1
                        if reportline == 32:
                            if report_version == "LARGE":
                                targetpage = "REGALIA_SALES_DTL_7b"
                                reportline = 20
                            else:
                                break

        # 11. Inventory.
        app.StatusBar = f"Inventory...{ledger_wb.Sheets(eql).Range('U7').Value}"
        if ledger_wb.Sheets(eql).Range("U7") > 0:
            msg_box(
                f"There are {ledger_wb.Sheets(eql).Range('U7').Value} Inventory Assets marked on the Equipment List.",
                VB_OK_ONLY,
                "",
            )
            if workbook.Sheets("INVENTORY_DTL_6").Range("E13") != "":
                msg = "Inventory information is already in this Report Form. Overwrite?"
                style = VB_YES_NO + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
                doitresponse = msg_box(msg, style, title)
            else:
                doitresponse = VB_YES
            if doitresponse == VB_YES:
                reportcol = 5
                targetpage = "INVENTORY_DTL_6"
                for ledgerline in range(11, 261):
                    if ledger_wb.Sheets(eql).Cells(ledgerline, 8).Value == "Inventory":
                        workbook.Sheets(targetpage).Cells(13, reportcol).Value = ledger_wb.Sheets(eql).Cells(
                            ledgerline, 4
                        ).Value
                        if app.WorksheetFunction.IsNumber(ledger_wb.Sheets(eql).Cells(ledgerline, 5)):
                            workbook.Sheets(targetpage).Cells(16, reportcol).Value = ledger_wb.Sheets(eql).Cells(
                                ledgerline, 5
                            ).Value
                        if ledger_wb.Sheets(eql).Cells(ledgerline, 6).Value < lyear:
                            workbook.Sheets(targetpage).Cells(17, reportcol).Value = ledger_wb.Sheets(eql).Cells(
                                ledgerline, 7
                            ).Value
                        else:
                            workbook.Sheets(targetpage).Cells(20, reportcol).Value = ledger_wb.Sheets(eql).Cells(
                                ledgerline, 7
                            ).Value
                        reportcol += 1
                        if reportcol == 13:
                            if report_version == "LARGE":
                                if targetpage == "INVENTORY_DTL_6":
                                    targetpage = "INVENTORY_DTL_6b"
                                reportcol = 5
                            else:
                                break

        # 12. Funds.
        app.StatusBar = "Funds..."
        workbook.Sheets("FUNDS_14").Range("F14:F55,D15:E55").ClearContents()
        workbook.Sheets("FUNDS_14").Range("f14").Value = ledger_wb.Sheets("Summary").Range("i10")
        if workbook.Sheets("FUNDS_14").Range("D15") != "":
            msg = "Fund information is already in this Report Form. Overwrite?"
            style = VB_YES_NO + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
            doitresponse = msg_box(msg, style, title)
        else:
            doitresponse = VB_YES
        if doitresponse == VB_YES:
            for i in range(10, 52):
                if ledger_wb.Sheets("Summary").Cells(i, 7).Value == "":
                    break
                if i > 10:
                    workbook.Sheets("FUNDS_14").Cells(i + 4, 4).Value = ledger_wb.Sheets("Summary").Cells(i, 7)
                fundtot = ledger_wb.Sheets("Summary").Cells(i, 8).Value
                if qtrarray[4]:
                    fundtot += (
                        ledger_wb.Sheets("Ledger_Q1").Range(f"AQ{i + 1}").Value
                        + ledger_wb.Sheets("Ledger_Q2").Range(f"AQ{i + 1}").Value
                        + ledger_wb.Sheets("Ledger_Q3").Range(f"AQ{i + 1}").Value
                        + ledger_wb.Sheets("Ledger_Q4").Range(f"AQ{i + 1}").Value
                    )
                elif qtrarray[3]:
                    fundtot += (
                        ledger_wb.Sheets("Ledger_Q1").Range(f"AQ{i + 1}").Value
                        + ledger_wb.Sheets("Ledger_Q2").Range(f"AQ{i + 1}").Value
                        + ledger_wb.Sheets("Ledger_Q3").Range(f"AQ{i + 1}").Value
                    )
                elif qtrarray[2]:
                    fundtot += (
                        ledger_wb.Sheets("Ledger_Q1").Range(f"AQ{i + 1}").Value
                        + ledger_wb.Sheets("Ledger_Q2").Range(f"AQ{i + 1}").Value
                    )
                elif qtrarray[1]:
                    fundtot += ledger_wb.Sheets("Ledger_Q1").Range(f"AQ{i + 1}").Value
                workbook.Sheets("FUNDS_14").Cells(i + 4, 6).Value = fundtot

    # 13. Transfers in.
    app.StatusBar = "Transfers In..."
    if getledgervalue(ledger_wb, "ax19", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR19").Value
        writereportline(ledger_wb, workbook, matchname, "TRANSFER_IN_9", 14, 39, 15, qtrarray)
    if getledgervalue(ledger_wb, "ax20", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR20").Value
        writereportline(ledger_wb, workbook, matchname, "TRANSFER_IN_9", 42, 57, 15, qtrarray)

    # 14. Transfers out.
    app.StatusBar = "Transfers Out..."
    if getledgervalue(ledger_wb, "ay49", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT49").Value
        writereportline(ledger_wb, workbook, matchname, "TRANSFER_OUT_10", 11, 24, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay50", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT50").Value
        writereportline(ledger_wb, workbook, matchname, "TRANSFER_OUT_10", 41, 50, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay51", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT51").Value
        writereportline(ledger_wb, workbook, matchname, "TRANSFER_OUT_10", 29, 38, 16, qtrarray)

    # 16. Income detail.
    app.StatusBar = "Income Detail..."
    if getledgervalue(ledger_wb, "ax11", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR11").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11a", 11, 20, 15, qtrarray)
    if getledgervalue(ledger_wb, "ax12", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR12").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11a", 23, 29, 15, qtrarray)
    if getledgervalue(ledger_wb, "ax13", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR13").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11a", 40, 51, 15, qtrarray)
    if getledgervalue(ledger_wb, "ax17", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR17").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11b", 12, 26, 15, qtrarray)
    if getledgervalue(ledger_wb, "ay52", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("At52").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11b", 12, 26, 16, qtrarray)
    if getledgervalue(ledger_wb, "ax18", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR18").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11b", 29, 35, 15, qtrarray)
    if getledgervalue(ledger_wb, "ax23", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR23").Value
        writereportline(ledger_wb, workbook, matchname, "REGALIA_SALES_DTL_7", 37, 46, 15, qtrarray)
    if getledgervalue(ledger_wb, "ax25", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR25").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11b", 40, 46, 15, qtrarray)
    if getledgervalue(ledger_wb, "ax26", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AR26").Value
        writereportline(ledger_wb, workbook, matchname, "INCOME_DTL_11b", 50, 56, 15, qtrarray)

    # 17. Expense detail.
    app.StatusBar = "Expense Detail..."
    if getledgervalue(ledger_wb, "ay11", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT11").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12a", 12, 22, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay12", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT12").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12a", 27, 38, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay13", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT13").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12a", 27, 38, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay14", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT14").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12a", 27, 38, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay21", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT21").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12a", 43, 54, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay22", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT22").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12a", 43, 54, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay23", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT23").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12a", 43, 54, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay30", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT30").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12b", 12, 21, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay47", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT47").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12b", 27, 41, 16, qtrarray)
    if getledgervalue(ledger_wb, "ay48", qtrarray) > 0:
        matchname = ledger_wb.Sheets("Ledger_Q1").Range("AT48").Value
        writereportline(ledger_wb, workbook, matchname, "EXPENSE_DTL_12b", 47, 55, 16, qtrarray)

    # Import quarter totals into income statement.
    _import_qtr(app, workbook, ledger_wb, "Ledger_Q1", qtrarray[1])
    _import_qtr(app, workbook, ledger_wb, "Ledger_Q2", qtrarray[2])
    _import_qtr(app, workbook, ledger_wb, "Ledger_Q3", qtrarray[3])
    _import_qtr(app, workbook, ledger_wb, "Ledger_Q4", qtrarray[4])

    # Save and close ledger workbook.
    app.StatusBar = f"Closing {ledgername}"
    ledger_wb.Saved = True
    ledger_wb.Close()
    workbook.Activate()

    app.StatusBar = f"Saving {workbook.Name}"
    bname = Module3.sanitize(workbook.Sheets("Contents").Range("C8"))
    new_name = (
        f"IMP_LDGR_{bname}_{workbook.Sheets('Contents').Range('C11')}_Q{workbook.Sheets('Contents').Range('C12')}"
    )
    app.StatusBar = f"Saving {new_name}"
    Module4.mysavefile(app, workbook, new_name)

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


def writereportline(
    ledger_wb: Any,
    report_wb: Any,
    matchtxt: Any,
    reportpage: str,
    rptlinestart: int,
    rptlineend: int,
    ledgercolstart: int,
    qtrarray: Sequence[bool],
) -> int:
    rptline = rptlinestart
    while rptline < rptlineend + 1:
        if matchtxt in {"Advert Non-SCA", "Insurance - NON SCA"}:
            break
        if report_wb.Sheets(reportpage).Cells(rptline, 3).Value != "":
            rptline += 1
        else:
            break

    if qtrarray[1]:
        rptline = writereportsub(
            ledger_wb,
            report_wb,
            rptline,
            "Ledger_Q1",
            matchtxt,
            reportpage,
            rptlineend,
            ledgercolstart,
        )
    if rptline < rptlineend and qtrarray[2]:
        rptline = writereportsub(
            ledger_wb,
            report_wb,
            rptline,
            "Ledger_Q2",
            matchtxt,
            reportpage,
            rptlineend,
            ledgercolstart,
        )
    if rptline < rptlineend and qtrarray[3]:
        rptline = writereportsub(
            ledger_wb,
            report_wb,
            rptline,
            "Ledger_Q3",
            matchtxt,
            reportpage,
            rptlineend,
            ledgercolstart,
        )
    if rptline < rptlineend and qtrarray[4]:
        rptline = writereportsub(
            ledger_wb,
            report_wb,
            rptline,
            "Ledger_Q4",
            matchtxt,
            reportpage,
            rptlineend,
            ledgercolstart,
        )
    return rptline


def writereportsub(
    ledger_wb: Any,
    report_wb: Any,
    rline: int,
    ledgerpage: str,
    matchtxt: Any,
    reportpage: str,
    rptlineend: int,
    ledgercolstart: int,
) -> int:
    max_lines = int(ledger_wb.Sheets(ledgerpage).Range("BR10").Value) + 10
    for ledgerline in range(11, max_lines + 1):
        if (
            ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart).Value == matchtxt
            or ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 5).Value == matchtxt
            or ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 11).Value == matchtxt
            or ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 16).Value == matchtxt
        ):
            ledgerfld1 = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 8).Value
            ledgerfld2 = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 9).Value
            ledgerfld3 = 0
            incflag = False
            if ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart).Value == matchtxt:
                ledgerfld3 = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 13).Value
                incflag = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 1).Value != ""
            if ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 5).Value == matchtxt:
                ledgerfld3 += ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 18).Value
                incflag = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 6).Value != ""
            if ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 11).Value == matchtxt:
                ledgerfld3 += ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 24).Value
                incflag = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 12).Value != ""
            if ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 16).Value == matchtxt:
                ledgerfld3 += ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 29).Value
                incflag = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, ledgercolstart + 17).Value != ""

            ledgerfld4 = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 4).Value
            ledgerfld5 = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 5).Value
            ledgerfld6 = ledger_wb.Sheets(ledgerpage).Cells(ledgerline, 10).Value

            if reportpage == "ASSET_DTL_5a":
                if not incflag:
                    ledgerfld3 *= -1
                if matchtxt == "Receivables":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2
                    report_wb.Sheets(reportpage).Cells(rline, 7).Value = ledgerfld3
                else:
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 7).Value = ledgerfld3
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") == "LARGE":
                        reportpage = "ASSET_DTL_5c"
                        if matchtxt == "Receivables":
                            rline = 13
                            rptlineend = 32
                        elif matchtxt == "Prepaid Expenses":
                            rline = 39
                            rptlineend = 43
                        elif matchtxt == "Other Assets":
                            rline = 50
                            rptlineend = 57
                    else:
                        break
            elif reportpage == "ASSET_DTL_5c":
                if not incflag:
                    ledgerfld3 *= -1
                if matchtxt == "Receivables":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                else:
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
            elif reportpage == "LIABILITY_DTL_5b":
                if not incflag:
                    ledgerfld3 *= -1
                if matchtxt == "Deferred Revenue":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                else:
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") == "LARGE":
                        reportpage = "LIABILITY_DTL_5d"
                        if matchtxt == "Deferred Revenue":
                            rline = 11
                            rptlineend = 28
                        elif matchtxt == "Payables":
                            rline = 33
                            rptlineend = 46
                        elif matchtxt == "Other Liabilities":
                            rline = 51
                            rptlineend = 55
                    else:
                        break
            elif reportpage == "LIABILITY_DTL_5d":
                if not incflag:
                    ledgerfld3 *= -1
                if matchtxt == "Deferred Revenue":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                else:
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
            elif reportpage == "REGALIA_SALES_DTL_7":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 9).Value = ledgerfld3
            elif reportpage == "INCOME_DTL_11a":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1
                report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld2
                report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld3
            elif reportpage == "INCOME_DTL_11b":
                if matchtxt == "Other Income":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                elif matchtxt == "Refund":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld3
                elif matchtxt == "Advertising Income":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld6
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld3
                else:
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld6
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld3
            elif reportpage == "EXPENSE_DTL_12a":
                if matchtxt == "Advert Non-SCA":
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                else:
                    xarray = str(matchtxt).split(" ")
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = xarray[-1]
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld1
                    report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld2
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
            elif reportpage == "EXPENSE_DTL_12b":
                if matchtxt == "Insurance - NON SCA":
                    report_wb.Sheets(reportpage).Cells(rline, 4).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                elif matchtxt == "Other Expense":
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = ledgerfld1
                    report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld2
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                else:
                    report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                    report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                    if rline == rptlineend:
                        if report_wb.Sheets("Contents").Range("C15") != "Corporate":
                            reportpage = "EXPENSE_DTL_12c"
                            rline = 11
                            rptlineend = 54
                        else:
                            break
            elif reportpage == "EXPENSE_DTL_12c":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld3
            elif reportpage == "TRANSFER_IN_9":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") != "SMALL":
                        reportpage = "TRANSFER_IN_9b"
                        if matchtxt == "Transfer In - In Kingdom":
                            rline = 11
                            rptlineend = 31
                        else:
                            rline = 36
                            rptlineend = 53
                    else:
                        break
            elif reportpage == "TRANSFER_IN_9b":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") != "MEDIUM":
                        reportpage = "TRANSFER_IN_9c"
                        if matchtxt == "Transfer In - In Kingdom":
                            rline = 11
                            rptlineend = 31
                        else:
                            rline = 36
                            rptlineend = 53
                    else:
                        break
            elif reportpage == "TRANSFER_IN_9c":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") != "MEDIUM" and matchtxt == "Transfer In - In Kingdom":
                        reportpage = "TRANSFER_IN_9d"
                        rline = 11
                        rptlineend = 54
                    else:
                        break
            elif reportpage == "TRANSFER_OUT_10":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld4
                report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld5
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") != "SMALL":
                        reportpage = "TRANSFER_OUT_10b"
                        if matchtxt == "Transfer Out - In Kingdom":
                            rline = 11
                            rptlineend = 27
                        elif matchtxt == "Transfer Out - SCA Corp":
                            rline = 32
                            rptlineend = 41
                        elif matchtxt == "Transfer Out - Out Kingdom":
                            rline = 44
                            rptlineend = 53
                    else:
                        break
            elif reportpage == "TRANSFER_OUT_10b":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld4
                report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld5
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") != "MEDIUM":
                        reportpage = "TRANSFER_OUT_10c"
                        if matchtxt == "Transfer Out - In Kingdom":
                            rline = 11
                            rptlineend = 27
                        elif matchtxt == "Transfer Out - SCA Corp":
                            rline = 32
                            rptlineend = 41
                        elif matchtxt == "Transfer Out - Out Kingdom":
                            rline = 44
                            rptlineend = 53
                    else:
                        break
            elif reportpage == "TRANSFER_OUT_10c":
                report_wb.Sheets(reportpage).Cells(rline, 3).Value = f"{ledgerfld1}, {ledgerfld2}"
                report_wb.Sheets(reportpage).Cells(rline, 6).Value = ledgerfld3
                report_wb.Sheets(reportpage).Cells(rline, 5).Value = ledgerfld4
                report_wb.Sheets(reportpage).Cells(rline, 4).Value = ledgerfld5
                if rline == rptlineend:
                    if report_wb.Sheets("CONTENTS").Range("B39") != "MEDIUM":
                        reportpage = "TRANSFER_OUT_10d"
                        if matchtxt == "Transfer Out - In Kingdom":
                            rline = 11
                            rptlineend = 27
                        elif matchtxt == "Transfer Out - SCA Corp":
                            rline = 32
                            rptlineend = 41
                        elif matchtxt == "Transfer Out - Out Kingdom":
                            rline = 44
                            rptlineend = 53
                    else:
                        break

            rline += 1
            if rline > rptlineend:
                msg_box(f"Too many {matchtxt} items - please fill in manually. Continuing...", VB_OK_ONLY, "")
                break
    return rline


def getledgervalue(ledger_wb: Any, rngname: str, qtrarray: Sequence[bool]) -> float:
    q1 = ledger_wb.Sheets("Ledger_Q1").Range(rngname).Value if qtrarray[1] else 0
    q2 = ledger_wb.Sheets("Ledger_Q2").Range(rngname).Value if qtrarray[2] else 0
    q3 = ledger_wb.Sheets("Ledger_Q3").Range(rngname).Value if qtrarray[3] else 0
    q4 = ledger_wb.Sheets("Ledger_Q4").Range(rngname).Value if qtrarray[4] else 0
    return q1 + q2 + q3 + q4


def getoutstandingchecks(
    ledger_wb: Any,
    report_wb: Any,
    reprow: int,
    repcol: int,
    qtrarray: Sequence[bool],
    months_array: List[str],
) -> None:
    months_array = list(months_array)
    months_array.extend(["N/A", "STALE"])

    for j in range(1, 5):
        do_it = False
        if j == 1 and qtrarray[1]:
            do_it = True
        elif j == 2 and qtrarray[2]:
            do_it = True
        elif j == 3 and qtrarray[3]:
            do_it = True
        elif j == 4 and qtrarray[4]:
            do_it = True
        ledgerpage = f"Ledger_Q{j}"
        if not do_it:
            continue
        for i in range(11, 111):
            if not Module3.inArray(months_array, ledger_wb.Sheets(ledgerpage).Cells(i, 7).Value):
                if ledger_wb.Sheets(ledgerpage).Cells(i, 6).Value < 0:
                    if report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol + 2).Value == "":
                        report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol).Value = ledger_wb.Sheets(
                            ledgerpage
                        ).Cells(i, 5).Value
                        report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol + 1).Value = ledger_wb.Sheets(
                            ledgerpage
                        ).Cells(i, 4).Value
                        report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol + 2).Value = (
                            ledger_wb.Sheets(ledgerpage).Cells(i, 6).Value * -1
                        )
                        reprow += 1
                    else:
                        if ledger_wb.Sheets(ledgerpage).Cells(i, 5).Value == "":
                            report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol).Value = (
                                f"{report_wb.Sheets('Primary_Account_2a').Cells(reprow, repcol).Value}, --"
                            )
                        else:
                            report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol).Value = (
                                f"{report_wb.Sheets('Primary_Account_2a').Cells(reprow, repcol).Value}, {ledger_wb.Sheets(ledgerpage).Cells(i, 5).Value}"
                            )
                        report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol + 2).Value = (
                            report_wb.Sheets("Primary_Account_2a").Cells(reprow, repcol + 2).Value
                            + (ledger_wb.Sheets(ledgerpage).Cells(i, 6).Value * -1)
                        )
                    if reprow == 35:
                        if repcol == 6:
                            reprow = 34
                        else:
                            reprow = 27
                            repcol += 3
        do_it = False
