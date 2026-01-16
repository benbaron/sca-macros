"""Translated cleanup and visibility routines from Module4.bas."""

from __future__ import annotations

from typing import Any

from .constants import BLACK, VB_DEFAULT_BUTTON1, VB_EXCLAMATION, VB_OK_ONLY
from .helpers import msg_box
from . import Module3

PWORD = "SCoE"
LPWORD = "KCoE"


def cleanupsub(app: Any, workbook: Any, nomsg: bool) -> None:
    app.DisplayAlerts = False
    app.StatusBar = "Resetting locks... "

    for sheet in workbook.Worksheets:
        sheet.Protect(PWORD)

    workbook.Sheets("FreeForm").Unprotect(PWORD)
    workbook.Sheets("Contents").Select()
    app.ScreenUpdating = True
    if not workbook.ProtectStructure:
        workbook.Protect(PWORD)

    mysavefile(app, workbook, workbook.Name)

    if not nomsg:
        msg = "Done! File Saved."
        msg_box(msg, VB_OK_ONLY + VB_EXCLAMATION + VB_DEFAULT_BUTTON1, "")

    workbook.Sheets("Contents").Select()
    app.ScreenUpdating = True
    app.DisplayAlerts = True
    app.StatusBar = False


def delete_unused(workbook: Any) -> None:
    for wks in workbook.Worksheets:
        wks.UsedRange
        try:
            last_row = wks.Cells.Find(
                "*",
                after=wks.Cells(1),
                LookIn=-4123,
                lookat=1,
                searchdirection=1,
                searchorder=1,
            ).Row
            last_col = wks.Cells.Find(
                "*",
                after=wks.Cells(1),
                LookIn=-4123,
                lookat=1,
                searchdirection=1,
                searchorder=2,
            ).Column
        except Exception:
            last_row = 0
            last_col = 0

        if last_row * last_col == 0:
            wks.Columns.Delete()
        else:
            wks.Range(wks.Cells(last_row + 1, 1), wks.Cells(wks.Rows.Count, 1)).EntireRow.Delete()
            wks.Range(wks.Cells(1, last_col + 1), wks.Cells(1, wks.Columns.Count)).EntireColumn.Delete()


def showstuff(workbook: Any) -> None:
    for sheet in workbook.Worksheets:
        sheet.Unprotect(PWORD)
        sheet.Columns("P:T").Hidden = False
    workbook.Sheets("Contents").Range("B38:B40").Locked = False
    workbook.Sheets("Contents").Rows("55:200").Hidden = False
    workbook.Sheets("Contents").Columns("H:Z").Hidden = False
    workbook.Sheets("Contents").Range("B37:B40").Font.Color = BLACK


def hidestuff(workbook: Any) -> None:
    for sheet in workbook.Worksheets:
        sheet.Unprotect(PWORD)
        if not (sheet.Name == "Contents" and sheet.Name == "Free Form"):
            if workbook.Sheets("Contents").Range("B39") == "LARGE":
                if workbook.Sheets("Contents").Range("C15") == "Corporate":
                    sheet.Columns("P").Hidden = False
                    sheet.Columns("Q:T").Hidden = True
                else:
                    sheet.Columns("T").Hidden = False
                    sheet.Columns("P:S").Hidden = True
            elif workbook.Sheets("Contents").Range("B39") == "MEDIUM":
                sheet.Columns("P").Hidden = True
                sheet.Columns("Q").Hidden = False
                sheet.Columns("R:T").Hidden = True
            elif workbook.Sheets("Contents").Range("B39") == "SMALL":
                sheet.Columns("P:Q").Hidden = True
                sheet.Columns("R").Hidden = False
                sheet.Columns("S:T").Hidden = True
            else:
                sheet.Columns("P:R").Hidden = True
                sheet.Columns("S").Hidden = False
                sheet.Columns("T").Hidden = True
        sheet.Select()
        sheet.Range("A1").Select()
        sheet.Parent.ActiveWindow.SmallScroll(Up=100, ToLeft=100)

    workbook.Sheets("Contents").Range("B38:B40").Locked = True
    workbook.Sheets("Contents").Rows("55:200").Hidden = True
    workbook.Sheets("Contents").Columns("H:Z").Hidden = True
    workbook.Sheets("Contents").Range("B37:B40").Font.Color = 0xCCFFCC

    for sheet in workbook.Worksheets:
        sheet.Protect(PWORD)
    workbook.Sheets("FreeForm").Unprotect(PWORD)


def mycopyfile(app: Any, workbook: Any, saveasname: str) -> None:
    app.DisplayAlerts = False
    app.StatusBar = f"Saving to new file {saveasname}"
    workbook.SaveCopyAs(f"{workbook.Path}\\{saveasname}")
    app.DisplayAlerts = True


def mysavefile(app: Any, workbook: Any, newsavename: str) -> None:
    app.DisplayAlerts = False
    saveasname = Module3.sanitize(newsavename)
    app.StatusBar = f"Saving to new file {saveasname}"
    try:
        workbook.SaveAs(f"{workbook.Path}\\{saveasname}")
    except Exception:
        saveOO(workbook, saveasname)
    app.DisplayAlerts = True


def saveOO(workbook: Any, saveasname: str) -> None:
    o_doc = workbook

    if o_doc.hasLocation:
        path = o_doc.getURL()
        ftype = path[-4:]
        while not path.endswith("/"):
            path = path[:-1]
    else:
        path = ""
        ftype = ".xls"

    if saveasname.endswith(".xls") or saveasname.endswith(".ods"):
        saveasname = saveasname[:-4]

    saveasname = f"{saveasname}{ftype}"

    prop = [
        {"Name": "Overwrite", "Value": True},
        {
            "Name": "FilterName",
            "Value": "MS Excel 97" if ftype == ".xls" else "StarOffice XML (Calc)",
        },
    ]

    o_doc.storeAsURL(f"{path}{saveasname}", prop)
