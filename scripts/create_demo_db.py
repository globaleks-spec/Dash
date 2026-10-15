"""Build a synthetic, clearly labelled database for cloud demonstrations."""

from __future__ import annotations

import sqlite3
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import column_index_from_string


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from dashboard.db import initialize, save_override  # noqa: E402
from dashboard.excel_import import import_workbook  # noqa: E402


GROUPS = [
    ("AN", "AS", "1 квартал 2026", ["БП 1 квартал", "Факт отгрузки", "Планируемая отгрузка", "Подписанные контракты", "Проекты акцепт Заказчика", "Прочие проекты в проработке"]),
    ("AT", "AZ", "2 квартал 2026", ["БП 2 квартал План", "БП-2 + ГЭП", "Факт отгрузки", "Планируемая отгрузка", "Подписанные контракты", "Проекты акцепт Заказчика", "Прочие проекты в проработке"]),
    ("BD", "BJ", "3 квартал 2026", ["БП 3 квартал План", "БП-3 + ГЭП", "Факт отгрузки", "Планируемая отгрузка", "Подписанные контракты", "Проекты акцепт Заказчика", "Прочие проекты в проработке"]),
    ("BN", "BT", "4 квартал 2026", ["БП 4 квартал План", "БП-4 + ГЭП", "Факт отгрузки", "Планируемая отгрузка", "Подписанные контракты", "Проекты акцепт Заказчика", "Прочие проекты в проработке"]),
    ("BX", "CD", "1 квартал 2027", ["БП 1 квартал", "Факт отгрузки", "Планируемая отгрузка", "Подписанные контракты", "Проекты акцепт Заказчика", "Прочие проекты в проработке", "Прочие проекты в проработке"]),
]


def make_book(path: Path, *, current: bool) -> Path:
    book = Workbook()
    sheet = book.active
    sheet.title = "Реализация БП 2026"
    sheet["C1"] = "Стоимость указана в т.р. без НДС · ТЕСТОВЫЕ ДАННЫЕ"
    sheet["A3"] = "Направление/ Менеджер"
    sheet["B3"] = "Контрагент"
    sheet["AM3"] = "БП 2026"
    for start, end, label, headers in GROUPS:
        sheet.merge_cells(f"{start}2:{end}2")
        sheet[f"{start}2"] = label
        start_col = column_index_from_string(start)
        for offset, header in enumerate(headers):
            sheet.cell(3, start_col + offset).value = header

    totals = [
        (4, "КОМПАНИЯ · ТЕСТ", 1_000_000, 250_000),
        (5, "Иванов И.И.", 400_000, 100_000),
        (6, "Петрова А.С.", 350_000, 90_000),
        (7, "Сидоров М.П.", 250_000, 60_000),
    ]
    for row, manager, annual, quarter in totals:
        sheet.cell(row, 1).value = manager
        sheet.cell(row, 2).value = "ИТОГО"
        sheet.cell(row, column_index_from_string("AM")).value = annual
        for column in ("AN", "AT", "BD", "BN"):
            sheet.cell(row, column_index_from_string(column)).value = quarter

    projects = [
        (9, "Иванов И.И.", "ТЕСТ · Заказчик Альфа · ИБП · отгружено 15.03.26", "AO", 120_000 if current else 100_000),
        (10, "Иванов И.И.", "ТЕСТ · Заказчик Бета · ИБП · производство", "BQ", 180_000 if current else 200_000),
        (11, "Петрова А.С.", "ТЕСТ · Заказчик Гамма · СЭЩ · ШОТ · подписанный договор", "BR", 140_000),
        (12, "Петрова А.С.", "ТЕСТ · Заказчик Дельта · СОПТ · акцепт", "BS", 160_000),
        (13, "Сидоров М.П.", "ТЕСТ · Заказчик Эпсилон · ЗВУ · участие в тендере", "BT", 220_000),
        (14, "Сидоров М.П.", "ТЕСТ · Заказчик Дзета · ИБП · перспектива 2027", "CC", 90_000),
    ]
    if current:
        projects.append((15, "Иванов И.И.", "ТЕСТ · Новый заказчик · выпрямитель · техническая проработка", "BT", 70_000))
    for row, manager, text, column, value in projects:
        sheet.cell(row, 1).value = manager
        sheet.cell(row, 2).value = text
        sheet.cell(row, column_index_from_string(column)).value = value

    data = book.create_sheet("Данные_new")
    data.append(["Менеджер СКД", "План", "Факт", "Прогноз", "% факт", "Дефицит", "Риск"])
    data.append(["ТЕСТОВЫЕ ДАННЫЕ", 1_000_000, 120_000 if current else 100_000, 700_000, 12, 300_000, 100_000])
    book.save(path)
    return path


def main() -> None:
    target = ROOT / "data" / "demo_dashboard.db"
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    connection = sqlite3.connect(target)
    connection.row_factory = sqlite3.Row
    initialize(connection)
    with tempfile.TemporaryDirectory(dir=ROOT) as temporary:
        temporary_path = Path(temporary)
        first = make_book(temporary_path / "ТЕСТ 05.10.26-09.10.26.xlsx", current=False)
        second = make_book(temporary_path / "ТЕСТ 12.10.26-16.10.26.xlsx", current=True)
        import_workbook(connection, first)
        result = import_workbook(connection, second)
    for source_row, fields in {
        11: {"probability": 0.9},
        12: {"probability": 0.7, "risk_schedule": 1},
        13: {"probability": 0.45, "risk_financial": 1},
        15: {"probability": 0.3, "risk_technical": 1},
    }.items():
        row = connection.execute(
            "SELECT business_key FROM project_snapshots WHERE import_id = ? AND source_row = ?",
            (result.import_id, source_row),
        ).fetchone()
        if row:
            save_override(connection, row["business_key"], fields)
    connection.close()
    print(f"Created synthetic demo database: {target}")


if __name__ == "__main__":
    main()
