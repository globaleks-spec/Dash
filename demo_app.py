"""Entrypoint for the private Streamlit Community Cloud demonstration."""

from __future__ import annotations

import os


os.environ["INVERTER_DB_PATH"] = "data/demo_dashboard.db"
os.environ["INVERTER_DEMO"] = "1"

# Importing app renders the Streamlit application.
import app  # noqa: E402,F401
