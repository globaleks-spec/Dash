"""Execute every Streamlit page in bare mode to catch runtime errors."""

from __future__ import annotations

import logging
import sys
import traceback
from pathlib import Path

logging.disable(logging.CRITICAL)
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app  # noqa: E402


def main() -> None:
    imports = app.list_imports(app.connection)
    if not imports:
        raise RuntimeError("Нет версии данных для проверки интерфейса")
    import_id = int(imports[0]["id"])
    year = 2026
    projects = app.deduplicate_projects(app.load_projects(app.connection, import_id))
    pages = [
        ("summary", lambda: app.page_summary(import_id, year, projects)),
        ("managers", lambda: app.page_managers(import_id, year, projects)),
        ("projects", lambda: app.page_projects(import_id, projects)),
        ("sesh", lambda: app.page_sesh(import_id, year, projects)),
        ("forecast", lambda: app.page_forecast(import_id, year, projects)),
        ("history", lambda: app.page_history(import_id, year)),
        ("quality", lambda: app.page_quality(import_id)),
        ("import", lambda: app.page_import(imports)),
        ("export", lambda: app.page_export(import_id, year)),
    ]
    failures: list[str] = []
    for name, callback in pages:
        try:
            callback()
        except Exception:
            failures.append(name)
            traceback.print_exc()
        else:
            print(f"OK {name}")
    if failures:
        raise SystemExit(f"UI smoke test failed: {', '.join(failures)}")


if __name__ == "__main__":
    main()
