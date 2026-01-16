"""Translated test data routines from Module6.bas."""

from __future__ import annotations

from typing import Any

from .constants import VB_DEFAULT_BUTTON1, VB_EXCLAMATION, VB_OK, VB_OK_ONLY, VB_YES, VB_YES_NO
from .helpers import msg_box
from . import Module4

PWORD = "SCoE"
LPWORD = "KCoE"


def ClearReport(app: Any, workbook: Any, clearme: bool, nomsg: bool) -> None:
    thisversion = workbook.Sheets("Contents").Range("B39")
    if workbook.Sheets("Contents").Range("B40") == "MASTER":
        thisversion = "MASTER"

    app.ScreenUpdating = False
    app.DisplayStatusBar = True

    msg = (
        "Do you want to save this cleared report workbook to a new file?"
        if clearme
        else "Do you want to save this messed up report workbook to a new file?"
    )
    style = VB_YES_NO + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    if msg_box(msg, style, "") == VB_YES:
        workbook.Sheets("Contents").Select()
        saveasname = (
            f"Report_{workbook.Sheets('Contents').Range('B39').Value}_"
            f"{workbook.Sheets('Contents').Range('B38').Value}"
        )
        Module4.mysavefile(app, workbook, saveasname)

    Module4.hidestuff(workbook)

    if clearme:
        workbook.Sheets("Contents").Range("C8:C11").ClearContents()
        workbook.Sheets("Contents").Range("C12") = 1
        workbook.Sheets("Contents").Range("C15") = "Corporate"
    else:
        sheet = workbook.Sheets("Contents")
        sheet.Range("C8") = "MESSED UP REPORT"
        sheet.Range("C9:C10") = "Someone Important"
        sheet.Range("C11") = 2009
        sheet.Range("C12") = 4
        sheet.Range("C15") = "Corporate"

    app.StatusBar = "Contact Info..."
    if clearme:
        with workbook.Sheets("CONTACT_INFO_1") as sheet:
            sheet.Range("D10:h10,D12:h12,D13:D14,F13,F14:h14,H13").ClearContents()
            sheet.Range("D15:f16,H15:H16,D18:h18,D19,f19,h19").ClearContents()
            sheet.Range("e21:h21,D22:h23,D24:D25,F24,F25:h25,H24").ClearContents()
            sheet.Range("D26:f27,H26:H27").ClearContents()
            sheet.Range("e29:H29,D30:h31,D32:d33,F32,F33:h33,H32").ClearContents()
            sheet.Range("D34:f35,H34:H35").ClearContents()
    else:
        with workbook.Sheets("CONTACT_INFO_1") as sheet:
            sheet.Range("D10,D12,D13:D14,F13,F14,H13,D15:D16,H15:H16") = 1
            sheet.Range("D18,D19,f19,h19") = 1
            sheet.Range("e21,D22,D23,D24:D25") = 1
            sheet.Range("F24,F25,H24,D26:D27,H26:H27") = 1
            sheet.Range("e29,D30,D31,D32:D33") = 1
            sheet.Range("F32,F33,H32,D34:D35,H34:H35") = 1

    app.StatusBar = "Primary Account..."
    if clearme:
        with workbook.Sheets("PRIMARY_ACCOUNT_2a") as sheet:
            sheet.Range("E13:h14,E15:E16,h16,F17:h17").ClearContents()
            sheet.Range("H15").Value = sheet.Range("C59").Value
            sheet.Range("h16,h19,C21:h23,C27:h34,h37,F38").ClearContents()
            sheet.Range("F38") = "No"
            sheet.Range("h40,C44:h53").ClearContents()
    else:
        with workbook.Sheets("PRIMARY_ACCOUNT_2a") as sheet:
            sheet.Range("E13:E14,E15:E16,h16,F17") = 1
            sheet.Range("H15").Value = sheet.Range("C59").Value
            sheet.Range("h16,h19,C21:h23,C27:h34,h37") = 1
            sheet.Range("F38") = "Yes"
            sheet.Range("h40,C44:h53") = 1

    app.StatusBar = "Secondary Accounts..."
    if clearme:
        workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range("D13:g21,D25:g25,D27:g44").ClearContents()
        workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range("D15:g15") = workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range(
            "C47"
        )
        if thisversion in {"LARGE", "MASTER"}:
            for name in ("SECONDARY_ACCOUNTS_2c", "SECONDARY_ACCOUNTS_2d"):
                workbook.Sheets(name).Range("D13:g21,D25:g25,D27:g44").ClearContents()
                workbook.Sheets(name).Range("D15:g15") = workbook.Sheets(name).Range("C47")
    else:
        workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range("D13:g21,D25:g25,D27:g44") = 1
        workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range("D16:g16") = "CD"
        workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range("D17:g17") = "No"
        workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range("D15:g15") = workbook.Sheets("SECONDARY_ACCOUNTS_2b").Range(
            "C47"
        )
        if thisversion in {"LARGE", "MASTER"}:
            for name in ("SECONDARY_ACCOUNTS_2c", "SECONDARY_ACCOUNTS_2d"):
                workbook.Sheets(name).Range("D13:g21,D25:g25,D27:g44") = 1
                workbook.Sheets(name).Range("D16:g16") = "CD"
                workbook.Sheets(name).Range("D17:g17") = "No"
                workbook.Sheets(name).Range("D15:g15") = workbook.Sheets(name).Range("C47")

    workbook.Sheets("BALANCE_3").Range("g19:g20").ClearContents()
    if clearme:
        if thisversion in {"MEDIUM", "LARGE", "MASTER"}:
            workbook.Sheets("BALANCE_3").Range("g31").ClearContents()
    else:
        workbook.Sheets("BALANCE_3").Range("g19:g20") = 1
        if thisversion in {"MEDIUM", "LARGE", "MASTER"}:
            workbook.Sheets("BALANCE_3").Range("g31") = 1

    app.StatusBar = "Cash Assets..."
    if clearme:
        workbook.Sheets("ASSET_DTL_5a").Range("c15:g18,c24:g34,c41:g45,c52:g59").ClearContents()
        if thisversion in {"LARGE", "PAYPAL", "MASTER"}:
            workbook.Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57").ClearContents()
        if thisversion in {"MEDIUM", "LARGE", "MASTER"}:
            app.StatusBar = "Non-cash Assets..."
            workbook.Sheets("INVENTORY_DTL_6").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents()
            workbook.Sheets("REGALIA_SALES_DTL_7").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents()
            workbook.Sheets("DEPR_DTL_8").Range("d14:g23,j14:j23,e32:g41,j32:j41").ClearContents()
            if thisversion in {"LARGE", "MASTER"}:
                workbook.Sheets("INVENTORY_DTL_6b").Range("E13:l14,E16:l17,E19:l20,E24:l25,E30:l30").ClearContents()
                workbook.Sheets("REGALIA_SALES_DTL_7b").Range("C20:H31,c37:I46,c49:g51,i49:I51").ClearContents()
                workbook.Sheets("DEPR_DTL_8b").Range("d14:g53,j14:j53").ClearContents()
                workbook.Sheets("DEPR_DTL_8c").Range("e14:g53,j14:j53").ClearContents()
    else:
        workbook.Sheets("ASSET_DTL_5a").Range("c15:g18,c24:g34,c41:g45,c52:g59") = 1
        if thisversion in {"LARGE", "PAYPAL", "MASTER"}:
            workbook.Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57") = 1
        if thisversion in {"MEDIUM", "LARGE", "MASTER"}:
            app.StatusBar = "Non-cash Assets..."
            with workbook.Sheets("INVENTORY_DTL_6") as sheet:
                sheet.Range("E14:l14") = 1
                sheet.Range("E16:l17") = 1
                sheet.Range("E19:l20") = 1
                sheet.Range("E24:l25") = 1
                sheet.Range("E30:l30") = 1
            with workbook.Sheets("REGALIA_SALES_DTL_7") as sheet:
                sheet.Range("C20:H31") = 1
                sheet.Range("c37:I46") = 1
                sheet.Range("c49:g51") = 1
                sheet.Range("i49:I51") = 1
            with workbook.Sheets("DEPR_DTL_8") as sheet:
                sheet.Range("d14:g23") = 1
                sheet.Range("j14:j23") = 1
                sheet.Range("e32:g41") = 1
                sheet.Range("j32:j41") = 1
            if thisversion in {"LARGE", "MASTER"}:
                with workbook.Sheets("INVENTORY_DTL_6b") as sheet:
                    sheet.Range("E14:l14") = 1
                    sheet.Range("E16:l17") = 1
                    sheet.Range("E19:l20") = 1
                    sheet.Range("E24:l25") = 1
                    sheet.Range("E30:l30") = 1
                with workbook.Sheets("REGALIA_SALES_DTL_7b") as sheet:
                    sheet.Range("C20:H31") = 1
                    sheet.Range("c37:I46") = 1
                    sheet.Range("c49:g51") = 1
                    sheet.Range("i49:I51") = 1
                workbook.Sheets("DEPR_DTL_8b").Range("d14:g53") = 1
                workbook.Sheets("DEPR_DTL_8b").Range("j14:j53") = 1
                workbook.Sheets("DEPR_DTL_8c").Range("e14:g53") = 1
                workbook.Sheets("DEPR_DTL_8c").Range("j14:j53") = 1

    MessIncomeExpense(workbook, thisversion) if not clearme else ClearIncomeExpense(workbook, thisversion)

    if not nomsg:
        Module4.cleanupsub(app, workbook, False)


def MessIncomeExpense(workbook: Any, thisvers: str) -> None:
    if thisvers in {"SMALL", "MEDIUM", "LARGE", "MASTER"}:
        workbook.Sheets("TRANSFER_IN_9").Range("c13:f57") = 1
        workbook.Sheets("TRANSFER_OUT_10").Range("c11:f50") = 1
        workbook.Sheets("INCOME_DTL_11a").Range("c11:e51") = 1
        workbook.Sheets("INCOME_DTL_11b").Range("c12:f56") = 1
        workbook.Sheets("EXPENSE_DTL_12a").Range("c12:f54") = 1
        workbook.Sheets("EXPENSE_DTL_12b").Range("c12:f55") = 1
        workbook.Sheets("FINANCE_COMM_13").Range("c11:f53") = 1
        workbook.Sheets("COMMENTS").Range("c8:c32") = "Testing"

    if thisvers in {"MEDIUM", "LARGE", "MASTER"}:
        workbook.Sheets("FUNDS_14").Range("d14:f55") = 1
        workbook.Sheets("NEWSLETTER_15").Range("d11:i58") = 1

    if thisvers in {"LARGE", "MASTER"}:
        workbook.Sheets("TRANSFER_IN_9b").Range("c11:f53") = 1
        workbook.Sheets("TRANSFER_IN_9c").Range("c11:f53") = 1
        workbook.Sheets("TRANSFER_IN_9d").Range("c11:f54") = 1
        workbook.Sheets("TRANSFER_OUT_10b").Range("c11:f53") = 1
        workbook.Sheets("TRANSFER_OUT_10c").Range("c11:f53") = 1
        workbook.Sheets("TRANSFER_OUT_10d").Range("c11:f53") = 1


def ClearIncomeExpense(workbook: Any, thisvers: str) -> None:
    workbook.Sheets("TRANSFER_IN_9").Range("c13:f57").ClearContents()
    workbook.Sheets("TRANSFER_OUT_10").Range("c11:f50").ClearContents()
    workbook.Sheets("INCOME_DTL_11a").Range("c11:e51").ClearContents()
    workbook.Sheets("INCOME_DTL_11b").Range("c12:f56").ClearContents()
    workbook.Sheets("EXPENSE_DTL_12a").Range("c12:f54").ClearContents()
    workbook.Sheets("EXPENSE_DTL_12b").Range("c12:f55").ClearContents()
    workbook.Sheets("FINANCE_COMM_13").Range("c11:f53").ClearContents()
    workbook.Sheets("COMMENTS").Range("c8:c32").ClearContents()

    if thisvers in {"MEDIUM", "LARGE", "MASTER"}:
        workbook.Sheets("FUNDS_14").Range("d14:f55").ClearContents()
        workbook.Sheets("NEWSLETTER_15").Range("d11:i58").ClearContents()

    if thisvers in {"LARGE", "MASTER"}:
        workbook.Sheets("TRANSFER_IN_9b").Range("c11:f53").ClearContents()
        workbook.Sheets("TRANSFER_IN_9c").Range("c11:f53").ClearContents()
        workbook.Sheets("TRANSFER_IN_9d").Range("c11:f54").ClearContents()
        workbook.Sheets("TRANSFER_OUT_10b").Range("c11:f53").ClearContents()
        workbook.Sheets("TRANSFER_OUT_10c").Range("c11:f53").ClearContents()
        workbook.Sheets("TRANSFER_OUT_10d").Range("c11:f53").ClearContents()


def messreport(app: Any, workbook: Any) -> None:
    ClearReport(app, workbook, False, False)


def unmessreport(app: Any, workbook: Any) -> None:
    ClearReport(app, workbook, True, False)
