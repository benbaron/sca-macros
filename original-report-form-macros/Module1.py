"""Printing utilities translated from the original VBA Module1.

This module centralizes report printing flows:
- counting and filtering printable pages from the Contents sheet
- printing in either forward or reverse order
- printing compact "four page" packs for quick review
"""

from __future__ import annotations

from typing import Any, List

from .constants import VB_DEFAULT_BUTTON1, VB_EXCLAMATION, VB_OK, VB_OK_CANCEL, VB_OK_ONLY
from .helpers import msg_box

PWORD = "SCoE"
LPWORD = "KCoE"


def print_backwards(app: Any, workbook: Any) -> None:
    """Print required report sheets in reverse order."""
    pages = fill_page()
    app.ScreenUpdating = False
    app.DisplayStatusBar = True
    app.StatusBar = "Printing Backwards..."

    count = 0
    # Identify required pages and mark non-required ones to skip.
    for idx in range(len(pages)):
        if workbook.Sheets("Contents").Cells(idx + 7, 7).Value == "REQUIRED":
            count += 1
        else:
            pages[idx] = "noprint"

    msg = f"You are about to print {count} pages."
    title = "Print Backwards"
    style = VB_OK_CANCEL + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    response = msg_box(msg, style, title)
    if response == VB_OK:
        # Iterate backwards through the list so sheets print in reverse order.
        for idx in range(len(pages) - 1, -1, -1):
            if pages[idx] != "noprint":
                workbook.Sheets(pages[idx]).Select()
                workbook.Sheets(pages[idx]).PrintOut()

    # Restore the UI state and notify the user.
    workbook.Sheets("Contents").Select()
    app.ScreenUpdating = True
    app.DisplayStatusBar = False
    msg_box("Done!", VB_OK_ONLY + VB_EXCLAMATION + VB_DEFAULT_BUTTON1, title)


def print_forwards(app: Any, workbook: Any) -> None:
    """Print required report sheets in forward order."""
    pages = fill_page()
    app.ScreenUpdating = False
    app.DisplayStatusBar = True
    app.StatusBar = "Printing Forwards..."

    count = 0
    # Identify required pages and mark non-required ones to skip.
    for idx in range(len(pages)):
        if workbook.Sheets("Contents").Cells(idx + 7, 7).Value == "REQUIRED":
            count += 1
        else:
            pages[idx] = "noprint"

    msg = f"You are about to print {count} pages."
    title = "Print Forwards"
    style = VB_OK_CANCEL + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    response = msg_box(msg, style, title)
    if response == VB_OK:
        # Iterate in normal order to print pages from front to back.
        for idx in range(len(pages)):
            if pages[idx] != "noprint":
                workbook.Sheets(pages[idx]).Select()
                workbook.Sheets(pages[idx]).PrintOut()

    # Restore the UI state and notify the user.
    workbook.Sheets("Contents").Select()
    app.ScreenUpdating = True
    app.DisplayStatusBar = False
    msg_box("Done!", VB_OK_ONLY + VB_EXCLAMATION + VB_DEFAULT_BUTTON1, title)


def print_four(app: Any, workbook: Any) -> None:
    """Print a condensed set of common report pages."""
    pages: List[str] = [
        "Contents",
        "CONTACT_INFO_1",
        "PRIMARY_ACCOUNT_2a",
        "SECONDARY_ACCOUNTS_2b",
        "SECONDARY_ACCOUNTS_2c",
        "SECONDARY_ACCOUNTS_2d",
        "BALANCE_3",
        "INCOME_4",
    ]
    count = 5
    app.ScreenUpdating = False
    app.DisplayStatusBar = True
    app.StatusBar = "Printing Forwards..."

    # Adjust secondary account pages based on form size and requirements.
    for idx in range(3, 6):
        if pages[idx] == "SECONDARY_ACCOUNTS_2c" and workbook.Sheets("Contents").Range("B39") == "LARGE":
            if workbook.Sheets("Contents").Range("G29").Value == "REQUIRED":
                count += 1
        elif pages[idx] == "SECONDARY_ACCOUNTS_2c":
            pages[idx] = "noprint"

        if pages[idx] == "SECONDARY_ACCOUNTS_2d" and workbook.Sheets("Contents").Range("B39") == "LARGE":
            if workbook.Sheets("Contents").Range("G30").Value == "REQUIRED":
                count += 1
        elif pages[idx] == "SECONDARY_ACCOUNTS_2d":
            pages[idx] = "noprint"

        if pages[idx] == "SECONDARY_ACCOUNTS_2b" and workbook.Sheets("Contents").Range("G29").Value == "REQUIRED":
            count += 1
        elif pages[idx] == "SECONDARY_ACCOUNTS_2b":
            pages[idx] = "noprint"

    msg = f"You are about to print {count} pages."
    title = "Print Report."
    style = VB_OK_CANCEL + VB_EXCLAMATION + VB_DEFAULT_BUTTON1
    response = msg_box(msg, style, title)
    if response == VB_OK:
        # Print the selected short list of pages.
        for idx in range(len(pages)):
            if pages[idx] != "noprint":
                workbook.Sheets(pages[idx]).Select()
                workbook.Sheets(pages[idx]).PrintOut()

    # Restore the UI state and notify the user.
    workbook.Sheets("Contents").Select()
    app.ScreenUpdating = True
    app.DisplayStatusBar = False
    msg_box("Done!", VB_OK_ONLY + VB_EXCLAMATION + VB_DEFAULT_BUTTON1, title)


def fill_page() -> List[str]:
    """Return the ordered list of report sheets used for printing."""
    return [
        "Contents",
        "CONTACT_INFO_1",
        "PRIMARY_ACCOUNT_2a",
        "SECONDARY_ACCOUNTS_2b",
        "BALANCE_3",
        "INCOME_4",
        "ASSET_DTL_5a",
        "LIABILITY_DTL_5b",
        "INVENTORY_DTL_6",
        "REGALIA_SALES_DTL_7",
        "DEPR_DTL_8",
        "TRANSFER_IN_9",
        "TRANSFER_OUT_10",
        "INCOME_DTL_11a",
        "INCOME_DTL_11b",
        "INCOME_DTL_11c",
        "EXPENSE_DTL_12a",
        "EXPENSE_DTL_12b",
        "FINANCE_COMM_13",
        "FUNDS_14",
        "NEWSLETTER_15",
        "COMMENTS",
        "UNUSED_LINE28",
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
        "EXPENSE_DTL_12c",
        "FreeForm",
    ]
