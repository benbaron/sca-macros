"""Translated report reset routine from Module5.bas."""

from __future__ import annotations

from typing import Any

from .constants import VB_DEFAULT_BUTTON1, VB_EXCLAMATION, VB_OK, VB_OK_CANCEL, VB_YES, VB_YES_NO
from .helpers import msg_box
from . import Module4, Module6

PWORD = "SCoE"
LPWORD = "KCoE"


def ResetReport(app: Any, workbook: Any) -> None:
    thisversion = workbook.Sheets("Contents").Range("B39")
    msg = "You are about to reset the entire report workbook for a new quarter!"
    title = "RESET Report"
    style = VB_OK_CANCEL + VB_EXCLAMATION + VB_DEFAULT_BUTTON1

    if msg_box(msg, style, title) != VB_OK:
        return

    app.ScreenUpdating = False
    app.DisplayStatusBar = True

    msg = "Do you want to save this reset report workbook to a new file?"
    style = VB_YES_NO + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    if msg_box(msg, style, title) == VB_YES:
        workbook.Sheets("Contents").Select()
        if app.WorksheetFunction.CountBlank(workbook.Sheets("Contents").Range("C8")) == 0:
            saveasname = f"Report_{workbook.Sheets('Contents').Range('C8').Value}_"
            if workbook.Sheets("Contents").Range("C12").Value == 4:
                saveasname += f"{workbook.Sheets('Contents').Range('C11').Value + 1}_Q1"
            else:
                saveasname += (
                    f"{workbook.Sheets('Contents').Range('C11').Value}_Q{workbook.Sheets('Contents').Range('C12').Value + 1}"
                )
        else:
            saveasname = f"New_{workbook.Name}"
        Module4.mysavefile(app, workbook, saveasname)

    app.StatusBar = "Resetting..."
    if workbook.Sheets("Contents").Range("C12").Value == 4:
        workbook.Sheets("Contents").Range("C11") = workbook.Sheets("Contents").Range("C11").Value + 1
        workbook.Sheets("Contents").Range("C12") = 1
    else:
        workbook.Sheets("Contents").Range("C12") = workbook.Sheets("Contents").Range("C12") + 1

    if workbook.Sheets("Contents").Range("C12").Value == 1 or workbook.Sheets("Contents").Range("C12") == "Sequential":
        workbook.Sheets("BALANCE_3").Range("g19:g20") = workbook.Sheets("BALANCE_3").Range("h19:h20").Value
        if thisversion in {"MEDIUM", "LARGE"}:
            workbook.Sheets("BALANCE_3").Range("g31") = workbook.Sheets("BALANCE_3").Range("h31").Value

    app.StatusBar = "Accounts..."
    with workbook.Sheets("PRIMARY_ACCOUNT_2a") as sheet:
        sheet.Range("h16").ClearContents()
        sheet.Range("h19").ClearContents()
        sheet.Range("C21:g23").ClearContents()
        sheet.Range("h37").ClearContents()

    with workbook.Sheets("SECONDARY_ACCOUNTS_2b") as sheet:
        sheet.Range("D18:g21").ClearContents()
        sheet.Range("D25:g25").ClearContents()

    if thisversion == "LARGE":
        for name in ("SECONDARY_ACCOUNTS_2c", "SECONDARY_ACCOUNTS_2d"):
            with workbook.Sheets(name) as sheet:
                sheet.Range("D18:g21").ClearContents()
                sheet.Range("D25:g25").ClearContents()

    if workbook.Sheets("Contents").Range("C12").Value == 1 or workbook.Sheets("Contents").Range("C13").Value == "Sequential":
        app.StatusBar = "Cash Assets..."
        with workbook.Sheets("ASSET_DTL_5a") as sheet:
            sheet.Range("c15:g18").ClearContents()
            sheet.Range("f24:f34") = sheet.Range("g24:g34").Value
            sheet.Range("f41:f45") = sheet.Range("g41:g45").Value
            sheet.Range("f52:f59") = sheet.Range("g52:g59").Value
            sheet.Range("g24:g34,g41:g45,g52:g59").ClearContents()

        if thisversion in {"LARGE", "PAYPAL"}:
            with workbook.Sheets("ASSET_DTL_5c") as sheet:
                sheet.Range("e13:e32") = sheet.Range("f13:f32").Value
                sheet.Range("f39:e43") = sheet.Range("f39:f43").Value
                sheet.Range("e50:e57") = sheet.Range("f50:f57").Value
                sheet.Range("f13:f32,f39:f43,f50:f57").ClearContents()

        if thisversion in {"MEDIUM", "LARGE"}:
            app.StatusBar = "Non-cash Assets..."
            with workbook.Sheets("INVENTORY_DTL_6") as sheet:
                sheet.Range("E16:L17") = sheet.Range("E26:L27")
                sheet.Range("E24:L25,E30:L30").ClearContents()
            with workbook.Sheets("REGALIA_SALES_DTL_7") as sheet:
                sheet.Range("f20:f31") = sheet.Range("I20:I31").Value
                sheet.Range("g20:h31").ClearContents()
                sheet.Range("c37:I46,c49:g51,i49:I51").ClearContents()
            if thisversion == "LARGE":
                with workbook.Sheets("INVENTORY_DTL_6b") as sheet:
                    sheet.Range("E16:L17") = sheet.Range("E26:L27")
                    sheet.Range("E24:L25,E30:L30").ClearContents()
                with workbook.Sheets("REGALIA_SALES_DTL_7b") as sheet:
                    sheet.Range("f20:f31") = sheet.Range("I20:I31").Value
                    sheet.Range("g20:h31").ClearContents()
                    sheet.Range("c37:I46,c49:g51,i49:I51").ClearContents()

        app.StatusBar = "Liabilities..."
        with workbook.Sheets("LIABILITY_DTL_5b") as sheet:
            sheet.Range("e16:e30") = sheet.Range("f16:f30").Value
            sheet.Range("e37:e43") = sheet.Range("f37:f43").Value
            sheet.Range("e49:e55") = sheet.Range("f49:f55").Value
            sheet.Range("f16:f30,f37:f42,f49:f55").ClearContents()

        if thisversion in {"LARGE", "PAYPAL"}:
            with workbook.Sheets("LIABILITY_DTL_5d") as sheet:
                sheet.Range("e11:e28") = sheet.Range("f11:f28").Value
                sheet.Range("e33:e46") = sheet.Range("f33:f46").Value
                sheet.Range("e51:e55") = sheet.Range("f51:f55").Value
                sheet.Range("f11:f28,f33:f46,f51:f55").ClearContents()
            if thisversion == "PAYPAL":
                for name in ("LIABILITY_DTL_5e", "LIABILITY_DTL_5f", "LIABILITY_DTL_5g"):
                    with workbook.Sheets(name) as sheet:
                        sheet.Range("e11:e55") = sheet.Range("f11:f55").Value
                        sheet.Range("f11:f55").ClearContents()

        app.StatusBar = "Newsletter Subscriptions..."
        if thisversion in {"MEDIUM", "LARGE"}:
            workbook.Sheets("NEWSLETTER_15").Range("I11,D22:E57,g22:h57,F58,I58").ClearContents()

        Module6.ClearIncomeExpense(workbook, thisversion)
    else:
        app.StatusBar = "Cash Assets..."
        workbook.Sheets("ASSET_DTL_5a").Range("c15:g18,g24:g34,g41:g45,g53:g59").ClearContents()
        if thisversion in {"LARGE", "PAYPAL"}:
            workbook.Sheets("ASSET_DTL_5c").Range("c13:f32,c39:f43,c50:f57").ClearContents()
        if thisversion in {"MEDIUM", "LARGE"}:
            app.StatusBar = "Clearing Ending Non-cash Asset Values..."
            workbook.Sheets("INVENTORY_DTL_6").Range("E24:L25,E30:L30").ClearContents()
            workbook.Sheets("REGALIA_SALES_DTL_7").Range("f20:f31") = workbook.Sheets("REGALIA_SALES_DTL_7").Range(
                "I20:I31"
            ).Value
            workbook.Sheets("REGALIA_SALES_DTL_7").Range("g20:h31").ClearContents()
            workbook.Sheets("REGALIA_SALES_DTL_7").Range("H37:I46").ClearContents()
            if thisversion == "LARGE":
                workbook.Sheets("INVENTORY_DTL_6b").Range("E24:L25,E30:L30").ClearContents()
                workbook.Sheets("REGALIA_SALES_DTL_7b").Range("f20:f31") = workbook.Sheets(
                    "REGALIA_SALES_DTL_7b"
                ).Range("I20:I31").Value
                workbook.Sheets("REGALIA_SALES_DTL_7b").Range("g20:h31").ClearContents()
                workbook.Sheets("REGALIA_SALES_DTL_7b").Range("H37:I46").ClearContents()

        app.StatusBar = "Liabilities..."
        workbook.Sheets("LIABILITY_DTL_5b").Range("f16:f30,f37:f43,f49:f55").ClearContents()
        if thisversion in {"LARGE", "PAYPAL"}:
            workbook.Sheets("LIABILITY_DTL_5d").Range("f11:f28,f33:f46,f51:f55").ClearContents()
            if thisversion == "PAYPAL":
                workbook.Sheets("LIABILITY_DTL_5e").Range("f11:f55").ClearContents()
                workbook.Sheets("LIABILITY_DTL_5f").Range("f11:f55").ClearContents()
                workbook.Sheets("LIABILITY_DTL_5g").Range("f11:f55").ClearContents()

    app.StatusBar = "Fund Balances..."
    if thisversion != "SMALL":
        workbook.Sheets("FUNDS_14").Range("F14:F55").ClearContents()

    app.StatusBar = "Comments..."
    workbook.Sheets("COMMENTS").Range("C8:C32").ClearContents()

    workbook.Sheets("FreeForm").Columns.Delete()

    Module4.cleanupsub(app, workbook, False)
