"""Utilities for importing Outlook .msg project information."""

from .msg_reader import OutlookMessage, parse_msg_file, OutlookMsgError
from .project_info_mapper import (
    ProjectInfoData,
    ProjectInfoParseResult,
    map_message_to_project_info,
)

__all__ = [
    "OutlookMsgError",
    "OutlookMessage",
    "parse_msg_file",
    "ProjectInfoData",
    "ProjectInfoParseResult",
    "map_message_to_project_info",
]
