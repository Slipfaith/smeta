"""Updater package."""
from .auto_updater import check_for_updates, check_for_updates_background
from .version import APP_VERSION, RELEASE_DATE, AUTHOR

__all__ = [
    "check_for_updates",
    "check_for_updates_background",
    "APP_VERSION",
    "RELEASE_DATE",
    "AUTHOR",
]
