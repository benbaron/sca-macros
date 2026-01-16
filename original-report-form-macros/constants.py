"""Shared constants used by the translated VBA macros.

These constants mirror the VBA/Excel values referenced throughout the
translated modules to keep behavior aligned with the original macros.
"""

from __future__ import annotations

# Standard MsgBox return codes.
VB_OK = 1
VB_CANCEL = 2
VB_ABORT = 3
VB_RETRY = 4
VB_IGNORE = 5
VB_YES = 6
VB_NO = 7

# Standard MsgBox button layouts.
VB_OK_ONLY = 0
VB_OK_CANCEL = 1
VB_ABORT_RETRY_IGNORE = 2
VB_YES_NO_CANCEL = 3
VB_YES_NO = 4
VB_RETRY_CANCEL = 5

# MsgBox icon/style flags.
VB_EXCLAMATION = 48
VB_DEFAULT_BUTTON1 = 0

# Excel calculation mode constants.
XL_MANUAL = -4135
XL_AUTOMATIC = -4105

# Common color values.
BLACK = 0x000000
