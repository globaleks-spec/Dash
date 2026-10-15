from __future__ import annotations

import argparse
import json
from pathlib import Path

from .db import DEFAULT_DB_PATH, connect, initialize
from .excel_import import import_workbook, profile_workbook, result_as_dict


def main() -> None:
    parser = argparse.ArgumentParser(description="Служебные команды дашборда продаж")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH, help="Путь к SQLite")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="Показать структуру Excel")
    inspect_parser.add_argument("path", type=Path)

    import_parser = subparsers.add_parser("import", help="Импортировать недельную версию")
    import_parser.add_argument("path", type=Path)

    args = parser.parse_args()
    connection = connect(args.db)
    initialize(connection)

    if args.command == "inspect":
        print(json.dumps(profile_workbook(args.path), ensure_ascii=False, indent=2))
    elif args.command == "import":
        result = import_workbook(connection, args.path)
        print(json.dumps(result_as_dict(result), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
