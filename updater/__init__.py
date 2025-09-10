"""Updater package."""
from .auto_updater import check_for_updates
from .version import APP_VERSION, RELEASE_DATE, AUTHOR

__all__ = ["check_for_updates", "APP_VERSION", "RELEASE_DATE", "AUTHOR"]
