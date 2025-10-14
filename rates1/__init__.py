"""Integration helpers for the legacy rates module."""

# Expose the main RateTab widget so it can be embedded into the
# application without importing from the internal path in multiple
# places.  This keeps the import paths short and mirrors how the
# standalone utility used these modules.

from .tabs.rate_tab import RateTab  # noqa: F401

