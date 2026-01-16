"""Helper utilities for the translated report form macros."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Iterable, Optional

from .constants import VB_OK


@dataclass
class MacroContext:
    """Minimal context wrapper for interacting with an Excel-like object model."""

    app: Any
    workbooks: Any
    active_workbook: Any

    @property
    def sheets(self) -> Any:
        return self.active_workbook.Sheets


MsgBoxHandler = Callable[[str, int, str], int]
_MSG_BOX_HANDLER: Optional[MsgBoxHandler] = None


def set_msg_box_handler(handler: Optional[MsgBoxHandler]) -> None:
    """Register a handler to resolve MsgBox prompts."""

    global _MSG_BOX_HANDLER
    _MSG_BOX_HANDLER = handler


def msg_box(message: str, style: int, title: str) -> int:
    """Placeholder MsgBox equivalent.

    Returns VB_OK by default so scripts can continue without UI.
    """

    if _MSG_BOX_HANDLER is not None:
        return _MSG_BOX_HANDLER(message, style, title)
    _ = (message, style, title)
    return VB_OK


def in_array(values: Iterable[Any], target: Any) -> bool:
    """Return True when `target` exists in the provided iterable."""
    return any(item == target for item in values)
