from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from dashboard.analytics import calculate_kpis, manager_summary, quality_issues
from dashboard.db import initialize, save_override
from dashboard.excel_import import import_workbook, profile_workbook
from tests.helpers import make_weekly_workbook


class DashboardTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(dir=Path.cwd())
        self.root = Path(self.temp.name)
        self.connection = sqlite3.connect(":memory:")
        self.connection.row_factory = sqlite3.Row
        initialize(self.connection)

    def tearDown(self) -> None:
        self.connection.close()
        self.temp.cleanup()

    def test_excel_structure_and_mapping(self) -> None:
        path = make_weekly_workbook(self.root / "Свод 12.10.26-16.10.26.xlsx")
        profile = profile_workbook(path)
        self.assertEqual(profile["primary_sheet"], "Реализация БП 2026")
        targets = {item["target"] for item in profile["suggested_mapping"]}
        self.assertIn("2026 Q4 · pipeline", targets)

    def test_kpi_calculation_deduplicates_quarter_carryover(self) -> None:
        path = make_weekly_workbook(self.root / "Свод 12.10.26-16.10.26.xlsx")
        result = import_workbook(self.connection, path)
        kpi = calculate_kpis(self.connection, result.import_id, 2026)
        self.assertEqual(kpi["plan"], 1_000_000)
        self.assertEqual(kpi["fact"], 100_000)
        self.assertEqual(kpi["pipeline"], 300_000)
        self.assertEqual(kpi["signed"], 50_000)
        self.assertEqual(kpi["base_forecast"], 100_000)
        self.assertEqual(kpi["optimistic"], 450_000)
        self.assertEqual(kpi["next_year"], 75_000)

    def test_probability_changes_base_forecast(self) -> None:
        path = make_weekly_workbook(self.root / "Свод 12.10.26-16.10.26.xlsx")
        result = import_workbook(self.connection, path)
        row = self.connection.execute(
            "SELECT business_key FROM project_snapshots WHERE source_row = 7"
        ).fetchone()
        save_override(self.connection, row["business_key"], {"probability": 0.5})
        kpi = calculate_kpis(self.connection, result.import_id, 2026)
        self.assertEqual(kpi["base_forecast"], 250_000)

    def test_manager_table_keeps_annual_kpis_when_quarter_is_selected(self) -> None:
        path = make_weekly_workbook(self.root / "Свод 12.10.26-16.10.26.xlsx")
        result = import_workbook(self.connection, path)
        summary = manager_summary(self.connection, result.import_id, 2026, quarter=4)
        manager = summary.loc[summary["Менеджер"] == "Иванов И.И."].iloc[0]
        self.assertEqual(manager["Годовой план"], 1_000_000)
        self.assertEqual(manager["Факт отгрузки"], 100_000)
        self.assertEqual(manager["Прогноз до конца года"], 100_000)
        self.assertEqual(manager["План отгрузок квартала"], 0)

    def test_repeated_file_is_idempotent_and_versions_are_preserved(self) -> None:
        first = make_weekly_workbook(self.root / "Свод 12.10.26-16.10.26.xlsx")
        result1 = import_workbook(self.connection, first)
        repeated = import_workbook(self.connection, first)
        self.assertTrue(repeated.already_imported)
        self.assertEqual(result1.import_id, repeated.import_id)

        second = make_weekly_workbook(self.root / "Свод 19.10.26-23.10.26.xlsx", changed=True)
        result2 = import_workbook(self.connection, second)
        self.assertNotEqual(result1.import_id, result2.import_id)
        count = self.connection.execute("SELECT COUNT(*) FROM imports").fetchone()[0]
        self.assertEqual(count, 2)

    def test_quality_control_flags_missing_dates_and_duplicates(self) -> None:
        path = make_weekly_workbook(self.root / "Свод 12.10.26-16.10.26.xlsx")
        result = import_workbook(self.connection, path)
        issues = quality_issues(self.connection, result.import_id)
        self.assertIn("Не указана точная дата отгрузки", set(issues["Проблема"]))
        self.assertIn("Неоднозначное совпадение или несколько стадий", set(issues["Проблема"]))


if __name__ == "__main__":
    unittest.main()
