"""Local version creation routines translated from the original VBA Module7.

This module creates a set of unlocked local report templates from the master
workbook and then applies final setup adjustments per file size/variant.
"""

from __future__ import annotations

from typing import Any

from .constants import VB_DEFAULT_BUTTON1, VB_EXCLAMATION, VB_OK, VB_OK_CANCEL
from .helpers import msg_box
from . import Module4, Module6

PWORD = "SCoE"
LPWORD = "KCoE"


def createlocalversions(app: Any, workbook: Any) -> None:
    """Create local template copies for each size and variant."""
    Module6.ClearReport(app, workbook, True, True)

    # Confirm before creating a large batch of local files.
    msg = "You are about to create all new local versions!"
    title = "CREATE Local Reports"
    style = VB_OK_CANCEL + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    if msg_box(msg, style, title) != VB_OK:
        Module4.cleanupsub(app, workbook, False)
        return

    app.ScreenUpdating = False
    app.DisplayStatusBar = True
    app.DisplayAlerts = False

    # Temporarily show hidden elements before copying.
    Module4.showstuff(workbook)

    workbook.Sheets("Contents").Select()
    workbook.Sheets("Contents").Range("B40") = "LOCAL"
    workbook.Sheets("Contents").Range("B38") = "unlocked"

    workbook.Sheets("Contents").Range("C15") = "Corporate"
    workbook.Sheets("Contents").Range("B39") = "LARGE"
    largecorp = f"SCAFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, largecorp)

    workbook.Sheets("Contents").Range("B39") = "MEDIUM"
    mediumcorp = f"SCAFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, mediumcorp)

    workbook.Sheets("Contents").Range("B39") = "SMALL"
    smallcorp = f"SCAFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, smallcorp)

    workbook.Sheets("Contents").Range("B39") = "PayPal"
    paypalcorp = f"SCAFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, paypalcorp)

    workbook.Sheets("Contents").Range("C15") = "Illinois"
    workbook.Sheets("Contents").Range("B39") = "LARGE"
    largesub = f"SCASubFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, largesub)

    workbook.Sheets("Contents").Range("B39") = "MEDIUM"
    mediumsub = f"SCASubFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, mediumsub)

    workbook.Sheets("Contents").Range("B39") = "SMALL"
    smallsub = f"SCASubFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, smallsub)

    workbook.Sheets("Contents").Range("B39") = "SMALL"
    workbook.Sheets("Contents").Range("C15") = "Non-US"
    smallxsub = f"SCAXUSFinancialReportv6_{workbook.Sheets('Contents').Range('B39').Value}_{workbook.Sheets('Contents').Range('B38').Value}.xlsm"
    Module4.mycopyfile(app, workbook, smallxsub)

    workbook.Sheets("Contents").Range("B39") = "LARGE"
    workbook.Sheets("Contents").Range("B40") = "MASTER"
    app.DisplayAlerts = True
    workbook.Save()

    # Apply post-processing tweaks to each newly created workbook.
    finishsetup(app, workbook, largecorp)
    finishsetup(app, workbook, mediumcorp)
    finishsetup(app, workbook, smallcorp)
    finishsetup(app, workbook, paypalcorp)
    finishsetup(app, workbook, largesub)
    finishsetup(app, workbook, mediumsub)
    finishsetup(app, workbook, smallsub)
    finishsetup(app, workbook, smallxsub)

    Module4.cleanupsub(app, workbook, False)


def finishsetup(app: Any, workbook: Any, workbookname: str) -> None:
    """Finalize each local workbook by locking ranges and updating formulas."""
    app.Workbooks.Open(f"{workbook.Path}\\{workbookname}")
    app.Workbooks(workbookname).Activate()

    # Ensure all sheets are unprotected for updates.
    for sheet in app.ActiveWorkbook.Worksheets:
        sheet.Unprotect(PWORD)

    app.StatusBar = f"Fix Table of Contents..{app.ActiveWorkbook.Name}"
    sheet = app.ActiveWorkbook.Sheets("Contents")
    sheet.Range("F7:H27").Locked = False
    sheet.Range("F30:H50").Locked = False

    if sheet.Range("B39").Value == "SMALL":
        # SMALL reports have fewer table-of-contents rows to show.
        sheet.Range("E15:H17").ClearContents()
        sheet.Range("E27:H27").ClearContents()
        sheet.Range("E30:H48").ClearContents()
        if sheet.Range("C15").Value == "Corporate":
            sheet.Range("E49:H49").ClearContents()
    elif sheet.Range("B39").Value == "MEDIUM":
        # MEDIUM reports have intermediate TOC settings.
        sheet.Range("E30:H43").ClearContents()
        sheet.Range("E45:H48").ClearContents()
        if sheet.Range("C15").Value == "Corporate":
            sheet.Range("E49:H49").ClearContents()
    elif sheet.Range("B39").Value == "LARGE":
        # LARGE reports keep most TOC entries.
        sheet.Range("E33:H38").ClearContents()
        if sheet.Range("C15").Value == "Corporate":
            sheet.Range("E49:H49").ClearContents()
    else:
        # PayPal or other variants require special TOC cleanup.
        sheet.Range("E15:H17").ClearContents()
        sheet.Range("E26:H27").ClearContents()
        sheet.Range("E30:H32").ClearContents()
        sheet.Range("E39:H49").ClearContents()
        sheet.Shapes.Range("B_ImportLedger").Delete()
        sheet.Shapes.Range("B_ImportReport").Delete()

    sheet.Range("F7:H27").Locked = True
    sheet.Range("F30:H50").Locked = True

    if sheet.Range("C15").Value == "Non-US":
        # Non-US variants use different date formats and disable ledger import.
        sheet.Range("C61") = "=IF(C59=\"\",\"\",TEXT(DATE(C63,C59,1),\"*dd/mm/yyyy\"))"
        sheet.Range("C62") = "=IF(C60=\"\",\"\",TEXT(DATE(C63,C60,C64),\"*dd/mm/yyyy\"))"
        sheet.Shapes.Range("B_ImportLedger").Delete()

    sheet.Select()
    oldtext = "LARGE"
    newtext = sheet.Range("B39").Value
    for hlink in sheet.Hyperlinks:
        if oldtext in hlink.Address:
            hlink.Address = hlink.Address.replace(oldtext, str(newtext))

    wb = app.ActiveWorkbook
    size = sheet.Range("B39").Value
    app.StatusBar = f"Fix Balance Statement..{wb.Name}"
    balance = wb.Sheets("BALANCE_3")
    if size == "SMALL":
        balance.Range("H19").Formula = (
            "=IF('PRIMARY_ACCOUNT_2a'!$F$38<>\"YES\",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$22+'ASSET_DTL_5a'!$G$19"
        )
        balance.Range("H20").Formula = (
            "=IF('PRIMARY_ACCOUNT_2a'!$F$38=\"YES\",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$23"
        )
        balance.Range("G21").Formula = "='ASSET_DTL_5a'!F35"
        balance.Range("H21").Formula = "='ASSET_DTL_5a'!G35"
        balance.Range("G22:H25").ClearContents()
        balance.Range("G26").Formula = "='ASSET_DTL_5a'!F46"
        balance.Range("H26").Formula = "='ASSET_DTL_5a'!G46"
        balance.Range("G27").Formula = "='ASSET_DTL_5a'!F60"
        balance.Range("H27").Formula = "='ASSET_DTL_5a'!G60"
        balance.Range("G32").Formula = "='LIABILITY_DTL_5b'!E31"
        balance.Range("H32").Formula = "='LIABILITY_DTL_5b'!F31"
        balance.Range("G33").Formula = "='LIABILITY_DTL_5b'!E44"
        balance.Range("H33").Formula = "='LIABILITY_DTL_5b'!F44"
        balance.Range("G34").Formula = "='LIABILITY_DTL_5b'!E56"
        balance.Range("H34").Formula = "='LIABILITY_DTL_5b'!F56"
        balance.Range("H31").ClearContents()
        balance.Range("g31").Interior.Color = balance.Range("g32").Interior.Color
        balance.Range("g31").Locked = True
    elif size == "MEDIUM":
        balance.Range("H19").Formula = (
            "=IF('PRIMARY_ACCOUNT_2a'!$F$38<>\"YES\",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$22+'ASSET_DTL_5a'!$G$19"
        )
        balance.Range("H20").Formula = (
            "=IF('PRIMARY_ACCOUNT_2a'!F38=\"YES\",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$I$23"
        )
        balance.Range("G21").Formula = "='ASSET_DTL_5a'!F35"
        balance.Range("H21").Formula = "='ASSET_DTL_5a'!G35"
        balance.Range("G22").Formula = "='INVENTORY_DTL_6'!M17"
        balance.Range("H22").Formula = "='INVENTORY_DTL_6'!M27"
        balance.Range("G23").Formula = "='REGALIA_SALES_DTL_7'!F32"
        balance.Range("H23").Formula = "='REGALIA_SALES_DTL_7'!I32"
        balance.Range("G24").Formula = (
            "='DEPR_DTL_8'!I47 + 'REGALIA_SALES_DTL_7'!F49 + 'REGALIA_SALES_DTL_7'!F50 + 'REGALIA_SALES_DTL_7'!F51"
        )
        balance.Range("H24").Formula = "='DEPR_DTL_8'!J47"
        balance.Range("G25").Formula = (
            "=-1*('DEPR_DTL_8'!K47+'REGALIA_SALES_DTL_7'!G49+'REGALIA_SALES_DTL_7'!G50+'REGALIA_SALES_DTL_7'!G51)"
        )
        balance.Range("H25").Formula = "=('DEPR_DTL_8'!M47*-1)"
        balance.Range("G26").Formula = "='ASSET_DTL_5a'!F46"
        balance.Range("H26").Formula = "='ASSET_DTL_5a'!G46"
        balance.Range("G27").Formula = "='ASSET_DTL_5a'!F60"
        balance.Range("H27").Formula = "='ASSET_DTL_5a'!G60"
        balance.Range("G32").Formula = "='LIABILITY_DTL_5b'!E31"
        balance.Range("H32").Formula = "='LIABILITY_DTL_5b'!F31"
        balance.Range("G33").Formula = "='LIABILITY_DTL_5b'!E44"
        balance.Range("H33").Formula = "='LIABILITY_DTL_5b'!F44"
        balance.Range("G34").Formula = "='LIABILITY_DTL_5b'!E56"
        balance.Range("H34").Formula = "='LIABILITY_DTL_5b'!F56"
    elif size == "LARGE":
        balance.Range("G33").Formula = "='LIABILITY_DTL_5b'!E44+'LIABILITY_DTL_5d'!E47"
        balance.Range("H33").Formula = "='LIABILITY_DTL_5b'!F44+'LIABILITY_DTL_5d'!F47"
    else:
        balance.Range("H19").Formula = (
            "=IF('PRIMARY_ACCOUNT_2a'!$F$38<>\"YES\",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$i$22+'ASSET_DTL_5a'!$G$19"
        )
        balance.Range("H20").Formula = (
            "=IF('PRIMARY_ACCOUNT_2a'!F38=\"YES\",IF('PRIMARY_ACCOUNT_2a'!$h$37='PRIMARY_ACCOUNT_2a'!$h$36,'PRIMARY_ACCOUNT_2a'!$h$37,0),0)+'SECONDARY_ACCOUNTS_2b'!$i$23"
        )
        balance.Range("G21").Formula = "='ASSET_DTL_5a'!F35"
        balance.Range("H21").Formula = "='ASSET_DTL_5a'!G35"
        balance.Range("G22:H25").ClearContents()
        balance.Range("G26").Formula = "='ASSET_DTL_5a'!F46"
        balance.Range("H26").Formula = "='ASSET_DTL_5a'!G46"
        balance.Range("G27").Formula = "='ASSET_DTL_5a'!F60"
        balance.Range("H27").Formula = "='ASSET_DTL_5a'!G60"
        balance.Range("H31").ClearContents()
        balance.Range("g31").Interior.Color = balance.Range("g32").Interior.Color
        balance.Range("g31").Locked = True

    app.StatusBar = f"Fix Income Statement..{wb.Name}"
    income = wb.Sheets("INCOME_4")
    if size == "SMALL":
        income.Range("j16").Formula = "='TRANSFER_IN_9'!F38"
        income.Range("j17").Formula = "='TRANSFER_IN_9'!F58"
        income.Range("H19:I19").ClearContents()
        income.Range("j20").ClearContents()
        income.Range("j21").ClearContents()
        income.Range("g30:I30").ClearContents()
        income.Range("h39").ClearContents()
        if sheet.Range("C15").Value == "Corporate":
            income.Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
        income.Range("j45").Formula = "='TRANSFER_OUT_10'!F25"
        income.Range("j46").Formula = "='TRANSFER_OUT_10'!F52"
    elif size == "MEDIUM":
        income.Range("j16").Formula = "='TRANSFER_IN_9'!F38+'TRANSFER_IN_9b'!F32"
        income.Range("j17").Formula = "='TRANSFER_IN_9'!F58+'TRANSFER_IN_9b'!F54"
        income.Range("h19").Formula = "='INVENTORY_DTL_6'!M30"
        income.Range("i19").Formula = "='INVENTORY_DTL_6'!M29"
        income.Range("j20").Formula = "='REGALIA_SALES_DTL_7'!I53"
        income.Range("g30").Formula = (
            "=SUMIF('DEPR_DTL_8'!$D14:$D23,\"OA\",'DEPR_DTL_8'!$L14:$L23)+SUMIF('DEPR_DTL_8'!$D32:$D41,\"OA\",'DEPR_DTL_8'!$L32:$L41)"
        )
        income.Range("h30").Formula = (
            "=SUMIF('DEPR_DTL_8'!$D14:$D23,\"AR\",'DEPR_DTL_8'!$L14:$L23)+SUMIF('DEPR_DTL_8'!$D32:$D41,\"AR\",'DEPR_DTL_8'!$L32:$L41)"
        )
        income.Range("i30").Formula = (
            "=SUMIF('DEPR_DTL_8'!$D14:$D23,\"FR\",'DEPR_DTL_8'!$L14:$L23)+SUMIF('DEPR_DTL_8'!$D32:$D41,\"FR\",'DEPR_DTL_8'!$L32:$L41)"
        )
        income.Range("h39").Formula = "='REGALIA_SALES_DTL_7'!H52"
        if sheet.Range("C15") == "Corporate":
            income.Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
        income.Range("j45").Formula = "='TRANSFER_OUT_10'!F25+'TRANSFER_OUT_10b'!F28"
        income.Range("j46").Formula = "='TRANSFER_OUT_10'!F52+'TRANSFER_OUT_10b'!F42+'TRANSFER_OUT_10b'!F54"
    elif size == "LARGE":
        if sheet.Range("C15") == "Corporate":
            income.Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
    else:
        income.Range("H19:I19").ClearContents()
        income.Range("j20").ClearContents()
        income.Range("j21").ClearContents()
        income.Range("g30:I30").ClearContents()
        income.Range("h39").ClearContents()
        income.Range("j44").Formula = "='EXPENSE_DTL_12b'!I56"
        income.Range("j45").Formula = "='TRANSFER_OUT_10'!F25+'TRANSFER_OUT_10b'!F28"
        income.Range("j46").Formula = "='TRANSFER_OUT_10'!F52+'TRANSFER_OUT_10b'!F42+'TRANSFER_OUT_10b'!F54"

    if size in {"SMALL", "PayPal"}:
        wb.Sheets("INCOME_DTL_11a").Range("C35") = ""
        wb.Sheets("INCOME_DTL_11a").Range("E35").ClearContents()
    elif size == "MEDIUM":
        wb.Sheets("INCOME_DTL_11a").Range("E35") = "='REGALIA_SALES_DTL_7'!H32"

    if sheet.Range("C15") == "Corporate":
        sheet.Range("C15").Locked = True
        sheet.Range("C15").Interior.Color = sheet.Range("B15").Interior.Color
        sheet.Range("C15").Validation.Delete()
    elif sheet.Range("C15") == "Non-US":
        sheet.Range("C15").Locked = True
        sheet.Range("C15").Interior.Color = sheet.Range("B15").Interior.Color
        sheet.Range("C15").Validation.Delete()
        wb.Sheets("EXPENSE_DTL_12b").Range("C46") = wb.Sheets("Corporations").Range("c1")
        wb.Sheets("EXPENSE_DTL_12b").Range("E46") = wb.Sheets("Corporations").Range("b1")
        wb.Sheets("EXPENSE_DTL_12b").Range("C46").Interior.Color = sheet.Range("B15").Interior.Color
        wb.Sheets("EXPENSE_DTL_12b").Range("E46").Interior.Color = sheet.Range("B15").Interior.Color
        wb.Sheets("INCOME_4").Range("D25") = "SCA, Inc. Stock Clerk expenses are General Supplies!"
        wb.Sheets("INCOME_4").Range("D25").Interior.ColorIndex = 40
        wb.Sheets("INCOME_DTL_11a").Range("C31") = "Transfers in from foreign branches (except PayPal) go under a) below!"
        wb.Sheets("INCOME_DTL_11a").Range("C31").Interior.ColorIndex = 40
        wb.Sheets("EXPENSE_DTL_12a").Range("D40") = "Transfers to SCA, Inc. for Insurance go here!"
        wb.Sheets("EXPENSE_DTL_12a").Range("D40").Interior.ColorIndex = 40
        wb.Sheets("EXPENSE_DTL_12b").Range("D43") = "Transfers to foreign branches and kingdom accounts go here!"
        wb.Sheets("EXPENSE_DTL_12b").Range("D43").Interior.ColorIndex = 40
    else:
        app.StatusBar = f"Fix State forms..{wb.Name}"
        j = 60
        k = 100
        for i in range(3, 60):
            if wb.Sheets("Corporations").Range(f"B{i}") == "":
                continue
            sheet.Range(f"B{k}") = wb.Sheets("Corporations").Range(f"A{i}")
            k += 1
            wb.Sheets("EXPENSE_DTL_12b").Range(f"C{j}") = wb.Sheets("Corporations").Range(f"c{i}")
            wb.Sheets("EXPENSE_DTL_12b").Range(f"E{j}") = wb.Sheets("Corporations").Range(f"b{i}")
            wb.Sheets("EXPENSE_DTL_12c").Range(f"C{j}") = wb.Sheets("Corporations").Range(f"c{i}")
            wb.Sheets("EXPENSE_DTL_12c").Range(f"E{j}") = wb.Sheets("Corporations").Range(f"b{i}")
            j += 1

        sheet.Range(f"B{k}") = wb.Sheets("Corporations").Range("A59")
        sheet.Range("C15").Validation.Modify(3, 1, 3, f"=$B$100:$B${k - 1}")
        wb.Sheets("INCOME_4").Range("D25") = "SCA, Inc. Stock Clerk expenses are General Supplies!"
        wb.Sheets("INCOME_4").Range("D25").Interior.ColorIndex = 40
        wb.Sheets("INCOME_DTL_11a").Range("C31") = (
            "Transfers in from out-of-state branches (except PayPal) go under a) below!"
        )
        wb.Sheets("INCOME_DTL_11a").Range("C31").Interior.ColorIndex = 40
        wb.Sheets("EXPENSE_DTL_12a").Range("D40") = "Transfers to SCA, Inc. for Insurance go here!"
        wb.Sheets("EXPENSE_DTL_12a").Range("D40").Interior.ColorIndex = 40
        wb.Sheets("EXPENSE_DTL_12b").Range("D43") = "Transfers to out-of-state branches and kingdom accounts go here!"
        wb.Sheets("EXPENSE_DTL_12b").Range("D43").Interior.ColorIndex = 40
        wb.Sheets("EXPENSE_DTL_12b").Range("C46") = wb.Sheets("Corporations").Range("c1")
        wb.Sheets("EXPENSE_DTL_12b").Range("E46") = wb.Sheets("Corporations").Range("b1")
        wb.Sheets("EXPENSE_DTL_12b").Range("C46").Interior.Color = sheet.Range("B15").Interior.Color
        wb.Sheets("EXPENSE_DTL_12b").Range("E46").Interior.Color = sheet.Range("B15").Interior.Color
        wb.Sheets("EXPENSE_DTL_12b").Range("C46:D46").Locked = True
        wb.Sheets("EXPENSE_DTL_12b").Range("E46").Locked = True
        wb.Sheets("TRANSFER_IN_9").Range("C12") = app.WorksheetFunction.Substitute(
            wb.Sheets("TRANSFER_IN_9").Range("C12"), "country", "state"
        )
        wb.Sheets("TRANSFER_IN_9").Range("C40") = app.WorksheetFunction.Substitute(
            wb.Sheets("TRANSFER_IN_9").Range("C40"), "country", "state"
        )
        wb.Sheets("TRANSFER_OUT_10").Range("C9").Formula = app.WorksheetFunction.Substitute(
            wb.Sheets("TRANSFER_OUT_10").Range("C9"), "country", "state"
        )
        wb.Sheets("TRANSFER_OUT_10").Range("C27").Formula = app.WorksheetFunction.Substitute(
            wb.Sheets("TRANSFER_OUT_10").Range("C27"), "country", "state"
        )
        wb.Sheets("TRANSFER_OUT_10").Rows(45).Copy()
        wb.Sheets("TRANSFER_OUT_10").Paste(Destination=wb.Sheets("TRANSFER_OUT_10").Rows(39))
        wb.Sheets("TRANSFER_OUT_10").Paste(Destination=wb.Sheets("TRANSFER_OUT_10").Rows(40))
        wb.Sheets("TRANSFER_OUT_10").Paste(Destination=wb.Sheets("TRANSFER_OUT_10").Rows(51))
        wb.Sheets("TRANSFER_OUT_10").Range("C52").Value = "TOTAL"
        wb.Sheets("TRANSFER_OUT_10").Range("f52").Value = "=sum(f29:f51)"

        if size != "SMALL":
            wb.Sheets("TRANSFER_IN_9b").Range("C9").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_IN_9b").Range("C9"), "country", "state"
            )
            wb.Sheets("TRANSFER_IN_9b").Range("C34").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_IN_9b").Range("C34"), "country", "state"
            )
            wb.Sheets("TRANSFER_OUT_10b").Range("C9").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_OUT_10b").Range("C9"), "country", "state"
            )
            wb.Sheets("TRANSFER_OUT_10b").Range("C30").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_OUT_10b").Range("C30"), "country", "state"
            )
            wb.Sheets("TRANSFER_OUT_10b").Rows(45).Copy()
            wb.Sheets("TRANSFER_OUT_10b").Paste(Destination=wb.Sheets("TRANSFER_OUT_10b").Rows(42))
            wb.Sheets("TRANSFER_OUT_10b").Paste(Destination=wb.Sheets("TRANSFER_OUT_10b").Rows(43))
            wb.Sheets("TRANSFER_OUT_10b").Range("e54").Value = "TOTAL"
            wb.Sheets("TRANSFER_OUT_10b").Range("f54").Value = "=sum(f32:f53)"

        if size in {"LARGE", "PAYPAL"}:
            wb.Sheets("TRANSFER_IN_9c").Range("C9").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_IN_9c").Range("C9"), "country", "state"
            )
            wb.Sheets("TRANSFER_IN_9c").Range("C34").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_IN_9c").Range("C340"), "country", "state"
            )
            wb.Sheets("TRANSFER_IN_9d").Range("C92").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_IN_9d").Range("C9"), "country", "state"
            )

        if size == "LARGE":
            wb.Sheets("TRANSFER_OUT_10c").Range("C9").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_OUT_10c").Range("C9"), "country", "state"
            )
            wb.Sheets("TRANSFER_OUT_10c").Range("C30").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_OUT_10c").Range("C30"), "country", "state"
            )
            wb.Sheets("TRANSFER_OUT_10c").Rows(45).Copy()
            wb.Sheets("TRANSFER_OUT_10c").Paste(Destination=wb.Sheets("TRANSFER_OUT_10c").Rows(42))
            wb.Sheets("TRANSFER_OUT_10c").Paste(Destination=wb.Sheets("TRANSFER_OUT_10c").Rows(43))
            wb.Sheets("TRANSFER_OUT_10c").Range("e54").Value = "TOTAL"
            wb.Sheets("TRANSFER_OUT_10c").Range("f54").Value = "=sum(f32:f53)"

            wb.Sheets("TRANSFER_OUT_10d").Range("C9").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_OUT_10d").Range("C9"), "country", "state"
            )
            wb.Sheets("TRANSFER_OUT_10d").Range("C30").Formula = app.WorksheetFunction.Substitute(
                wb.Sheets("TRANSFER_OUT_10d").Range("C30"), "country", "state"
            )
            wb.Sheets("TRANSFER_OUT_10d").Rows(45).Copy()
            wb.Sheets("TRANSFER_OUT_10d").Paste(Destination=wb.Sheets("TRANSFER_OUT_10d").Rows(42))
            wb.Sheets("TRANSFER_OUT_10d").Paste(Destination=wb.Sheets("TRANSFER_OUT_10d").Rows(43))
            wb.Sheets("TRANSFER_OUT_10d").Range("e54").Value = "TOTAL"
            wb.Sheets("TRANSFER_OUT_10d").Range("f54").Value = "=sum(f32:f53)"

    wb.Unprotect(PWORD)
    app.StatusBar = "Remove extra pages..."
    app.DisplayAlerts = False
    wb.Sheets("Corporations").Delete()
    if size == "LARGE":
        for name in (
            "LIABILITY_DTL_5e",
            "LIABILITY_DTL_5f",
            "LIABILITY_DTL_5g",
            "LIABILITY_DTL_5h",
            "LIABILITY_DTL_5i",
        ):
            wb.Sheets(name).Delete()
        if sheet.Range("C15") == "Corporate":
            wb.Sheets("EXPENSE_DTL_12c").Delete()
    elif size == "MEDIUM":
        for name in (
            "SECONDARY_ACCOUNTS_2c",
            "SECONDARY_ACCOUNTS_2d",
            "ASSET_DTL_5c",
            "LIABILITY_DTL_5d",
            "LIABILITY_DTL_5e",
            "LIABILITY_DTL_5f",
            "LIABILITY_DTL_5g",
            "LIABILITY_DTL_5h",
            "LIABILITY_DTL_5i",
            "INVENTORY_DTL_6b",
            "REGALIA_SALES_DTL_7b",
            "DEPR_DTL_8b",
            "DEPR_DTL_8c",
            "TRANSFER_IN_9c",
            "TRANSFER_IN_9d",
            "TRANSFER_OUT_10c",
            "TRANSFER_OUT_10d",
        ):
            wb.Sheets(name).Delete()
        if sheet.Range("C15") == "Corporate":
            wb.Sheets("EXPENSE_DTL_12c").Delete()
    elif size == "PayPal":
        for name in (
            "INVENTORY_DTL_6",
            "REGALIA_SALES_DTL_7",
            "DEPR_DTL_8",
            "NEWSLETTER_15",
            "SECONDARY_ACCOUNTS_2c",
            "SECONDARY_ACCOUNTS_2d",
            "INVENTORY_DTL_6b",
            "REGALIA_SALES_DTL_7b",
            "DEPR_DTL_8b",
            "DEPR_DTL_8c",
            "TRANSFER_OUT_10c",
            "TRANSFER_OUT_10d",
            "EXPENSE_DTL_12c",
        ):
            wb.Sheets(name).Delete()
    else:
        for name in (
            "INVENTORY_DTL_6",
            "REGALIA_SALES_DTL_7",
            "DEPR_DTL_8",
            "NEWSLETTER_15",
            "SECONDARY_ACCOUNTS_2c",
            "SECONDARY_ACCOUNTS_2d",
            "ASSET_DTL_5c",
            "LIABILITY_DTL_5d",
            "LIABILITY_DTL_5e",
            "LIABILITY_DTL_5f",
            "LIABILITY_DTL_5g",
            "LIABILITY_DTL_5h",
            "LIABILITY_DTL_5i",
            "INVENTORY_DTL_6b",
            "REGALIA_SALES_DTL_7b",
            "DEPR_DTL_8b",
            "DEPR_DTL_8c",
            "TRANSFER_IN_9b",
            "TRANSFER_IN_9c",
            "TRANSFER_IN_9d",
            "TRANSFER_OUT_10b",
            "TRANSFER_OUT_10c",
            "TRANSFER_OUT_10d",
        ):
            wb.Sheets(name).Delete()
