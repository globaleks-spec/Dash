from __future__ import annotations

import io
import sqlite3
from datetime import date
from typing import Any, Iterable

import pandas as pd

from .db import OVERRIDE_FIELDS, latest_import_id, query_all


OUTSTANDING_METRICS = ("confirmed", "signed", "accepted", "pipeline")
METRIC_PRIORITY = {"confirmed": 4, "signed": 3, "accepted": 2, "pipeline": 1}
INACTIVE_STAGES = {"проиграно", "отменено"}


def load_projects(connection: sqlite3.Connection, import_id: int | None = None) -> pd.DataFrame:
    import_id = import_id or latest_import_id(connection)
    if import_id is None:
        return pd.DataFrame()
    projects = pd.read_sql_query(
        "SELECT * FROM project_snapshots WHERE import_id = ?", connection, params=(import_id,)
    )
    if projects.empty:
        return projects
    overrides = pd.read_sql_query("SELECT * FROM project_overrides", connection)
    if not overrides.empty:
        merged = projects.merge(overrides, on="business_key", how="left", suffixes=("", "__override"))
        for field in OVERRIDE_FIELDS:
            override = f"{field}__override"
            if override in merged:
                merged[field] = merged[override].where(
                    merged[override].notna(), merged[field]
                )
        projects = merged[[column for column in projects.columns]].copy()
    projects["probability"] = pd.to_numeric(projects["probability"], errors="coerce")
    return projects


def deduplicate_projects(projects: pd.DataFrame) -> pd.DataFrame:
    if projects.empty:
        return projects.copy()
    frame = projects.copy()
    scored_columns = [
        "manager",
        "customer",
        "project_name",
        "channel",
        "equipment_type",
        "project_amount",
        "stage",
        "planned_ship_date",
        "actual_ship_date",
    ]
    frame["_completeness"] = frame[scored_columns].notna().sum(axis=1)
    frame["_text_length"] = frame["source_text"].fillna("").str.len()
    frame["duplicate_count"] = frame.groupby("business_key")["business_key"].transform("size")
    frame = frame.sort_values(
        ["business_key", "_completeness", "_text_length", "source_row"],
        ascending=[True, False, False, True],
    )
    return frame.drop_duplicates("business_key", keep="first").drop(
        columns=["_completeness", "_text_length"]
    )


def apply_filters(projects: pd.DataFrame, filters: dict[str, Any] | None) -> pd.DataFrame:
    if projects.empty or not filters:
        return projects.copy()
    frame = projects.copy()
    for field in ("manager", "customer", "industry", "channel", "equipment_type", "stage"):
        selected = filters.get(field)
        if selected:
            selected_values = selected if isinstance(selected, (list, tuple, set)) else [selected]
            frame = frame[frame[field].fillna("Не указано").isin(selected_values)]
    search = str(filters.get("search") or "").strip()
    if search:
        mask = (
            frame["project_name"].fillna("").str.contains(search, case=False, regex=False)
            | frame["customer"].fillna("").str.contains(search, case=False, regex=False)
            | frame["source_text"].fillna("").str.contains(search, case=False, regex=False)
        )
        frame = frame[mask]
    month = filters.get("month")
    if month:
        date_fields = ["planned_ship_date", "actual_ship_date", "contract_date"]
        month_mask = pd.Series(False, index=frame.index)
        for field in date_fields:
            parsed = pd.to_datetime(frame[field], errors="coerce")
            month_mask |= parsed.dt.month.eq(int(month))
        frame = frame[month_mask]
    return frame


def load_period_values(
    connection: sqlite3.Connection, import_id: int, business_keys: Iterable[str] | None = None
) -> pd.DataFrame:
    frame = pd.read_sql_query(
        "SELECT * FROM period_values WHERE import_id = ?", connection, params=(import_id,)
    )
    if business_keys is not None and not frame.empty:
        keys = set(business_keys)
        frame = frame[frame["business_key"].isin(keys)]
    return frame


def _dedup_period_values(values: pd.DataFrame) -> pd.DataFrame:
    if values.empty:
        return values.copy()
    return (
        values.groupby(["business_key", "year", "quarter", "metric"], as_index=False)["amount"]
        .max()
    )


def actual_by_project(values: pd.DataFrame, year: int, quarter: int | None = None) -> pd.Series:
    if values.empty:
        return pd.Series(dtype=float)
    current = _dedup_period_values(values)
    current = current[(current["year"] == year) & (current["metric"] == "actual")]
    if quarter:
        current = current[current["quarter"] == quarter]
    return current.groupby("business_key")["amount"].sum()


def current_outstanding(
    values: pd.DataFrame, year: int, quarter: int | None = None
) -> pd.DataFrame:
    if values.empty:
        return pd.DataFrame(columns=["business_key", "quarter", "metric", "amount"])
    frame = _dedup_period_values(values)
    frame = frame[(frame["year"] == year) & frame["metric"].isin(OUTSTANDING_METRICS)]
    frame = frame[frame["amount"].abs() > 1e-9]
    if quarter:
        frame = frame[frame["quarter"] == quarter]
    if frame.empty:
        return pd.DataFrame(columns=["business_key", "quarter", "metric", "amount"])
    latest = frame.groupby("business_key")["quarter"].max().rename("latest_quarter")
    frame = frame.merge(latest, on="business_key")
    frame = frame[frame["quarter"] == frame["latest_quarter"]].copy()
    frame["priority"] = frame["metric"].map(METRIC_PRIORITY)
    frame = frame.sort_values(
        ["business_key", "priority", "amount"], ascending=[True, False, False]
    )
    return frame.drop_duplicates("business_key", keep="first")[[
        "business_key", "quarter", "metric", "amount"
    ]]


def _plan_amount(
    connection: sqlite3.Connection,
    import_id: int,
    year: int,
    managers: list[str] | None = None,
    quarter: int | None = None,
) -> float:
    params: list[Any] = [import_id, year]
    if quarter:
        metric = "plan"
        sql = (
            "SELECT manager, MAX(amount) AS amount FROM manager_plans "
            "WHERE import_id = ? AND year = ? AND quarter = ? AND metric = ?"
        )
        params.extend([quarter, metric])
    else:
        sql = (
            "SELECT manager, MAX(amount) AS amount FROM manager_plans "
            "WHERE import_id = ? AND year = ? AND quarter IS NULL AND metric = 'annual_plan'"
        )
    if managers:
        placeholders = ",".join("?" for _ in managers)
        sql += f" AND scope = 'manager' AND manager IN ({placeholders})"
        params.extend(managers)
    else:
        sql += " AND scope = 'company' AND manager = '__TOTAL__'"
    sql += " GROUP BY manager"
    rows = query_all(connection, sql, params)
    if rows:
        return float(sum(row["amount"] for row in rows))
    if not managers:
        fallback_sql = sql.replace(
            "scope = 'company' AND manager = '__TOTAL__'", "scope = 'manager'"
        )
        fallback_rows = query_all(connection, fallback_sql, params)
        return float(sum(row["amount"] for row in fallback_rows))
    return 0.0


def calculate_kpis(
    connection: sqlite3.Connection,
    import_id: int,
    year: int,
    filters: dict[str, Any] | None = None,
    quarter: int | None = None,
) -> dict[str, float]:
    projects = deduplicate_projects(load_projects(connection, import_id))
    projects = apply_filters(projects, filters)
    values = load_period_values(connection, import_id, projects["business_key"] if not projects.empty else [])

    actual = actual_by_project(values, year, quarter)
    outstanding = current_outstanding(values, year, quarter)
    active_keys = set(
        projects.loc[~projects["stage"].isin(INACTIVE_STAGES), "business_key"].tolist()
    )
    if not outstanding.empty:
        outstanding = outstanding[outstanding["business_key"].isin(active_keys)]
        outstanding = outstanding.merge(
            projects[["business_key", "probability"]], on="business_key", how="left"
        )
    metric_totals = (
        outstanding.groupby("metric")["amount"].sum().to_dict()
        if not outstanding.empty
        else {}
    )
    confirmed = float(metric_totals.get("confirmed", 0.0))
    signed = float(metric_totals.get("signed", 0.0))
    accepted = float(metric_totals.get("accepted", 0.0))
    pipeline = float(metric_totals.get("pipeline", 0.0))
    fact = float(actual.sum()) if not actual.empty else 0.0

    weighted = 0.0
    if not outstanding.empty:
        risk_pool = outstanding[outstanding["metric"].isin(["signed", "accepted", "pipeline"])]
        weighted = float((risk_pool["amount"] * risk_pool["probability"].fillna(0.0)).sum())

    managers = None
    if filters and filters.get("manager"):
        selected = filters["manager"]
        managers = list(selected) if isinstance(selected, (list, tuple, set)) else [selected]
    plan = _plan_amount(connection, import_id, year, managers, quarter)
    conservative = fact + confirmed
    base = conservative + weighted
    optimistic = fact + confirmed + signed + accepted + pipeline

    next_year_outstanding = current_outstanding(values, year + 1)
    if not next_year_outstanding.empty:
        next_year_outstanding = next_year_outstanding[
            next_year_outstanding["business_key"].isin(active_keys)
        ]
    next_year = float(next_year_outstanding["amount"].sum()) if not next_year_outstanding.empty else 0.0

    return {
        "plan": plan,
        "fact": fact,
        "completion": fact / plan if plan else 0.0,
        "gap": max(plan - fact, 0.0),
        "confirmed": confirmed,
        "base_forecast": base,
        "conservative": conservative,
        "optimistic": optimistic,
        "signed": signed,
        "planned_contracts": accepted + pipeline,
        "accepted": accepted,
        "pipeline": pipeline,
        "next_year": next_year,
        "probability_known_share": (
            float(outstanding["probability"].notna().mean()) if not outstanding.empty else 0.0
        ),
    }


def manager_summary(
    connection: sqlite3.Connection, import_id: int, year: int, quarter: int | None = None
) -> pd.DataFrame:
    projects = deduplicate_projects(load_projects(connection, import_id))
    if projects.empty:
        return pd.DataFrame()
    rows: list[dict[str, Any]] = []
    for manager in sorted(projects["manager"].dropna().unique()):
        manager_projects = projects[projects["manager"] == manager]
        # The comparison table is annual by definition.  The selected quarter
        # affects only the dedicated shipment-plan column; otherwise choosing
        # Q4 would silently turn the annual plan, YTD fact and year-end forecast
        # into Q4-only figures.
        kpis = calculate_kpis(
            connection, import_id, year, {"manager": [manager]}
        )
        quarter_kpis = (
            calculate_kpis(
                connection,
                import_id,
                year,
                {"manager": [manager]},
                quarter=quarter,
            )
            if quarter
            else kpis
        )
        rows.append(
            {
                "Менеджер": manager,
                "Годовой план": kpis["plan"],
                "Факт отгрузки": kpis["fact"],
                "% выполнения": kpis["completion"],
                "Подписанные контракты": kpis["signed"],
                "План отгрузок квартала": quarter_kpis["confirmed"],
                "Прогноз до конца года": kpis["base_forecast"],
                "План заключения договоров": kpis["planned_contracts"],
                "Портфель следующего года": kpis["next_year"],
                "Проектов": int(len(manager_projects)),
            }
        )
    return pd.DataFrame(rows).sort_values(
        ["Факт отгрузки", "% выполнения"], ascending=False
    )


def quarterly_cumulative(
    connection: sqlite3.Connection,
    import_id: int,
    year: int,
    filters: dict[str, Any] | None = None,
) -> pd.DataFrame:
    projects = deduplicate_projects(load_projects(connection, import_id))
    projects = apply_filters(projects, filters)
    values = load_period_values(connection, import_id, projects["business_key"] if not projects.empty else [])
    actuals = _dedup_period_values(values)
    actuals = actuals[(actuals["year"] == year) & (actuals["metric"] == "actual")]
    actual_by_quarter = actuals.groupby("quarter")["amount"].sum().to_dict()
    managers = None
    if filters and filters.get("manager"):
        selected = filters["manager"]
        managers = list(selected) if isinstance(selected, (list, tuple, set)) else [selected]
    rows = []
    cumulative_plan = 0.0
    cumulative_fact = 0.0
    for quarter in range(1, 5):
        cumulative_plan += _plan_amount(connection, import_id, year, managers, quarter)
        cumulative_fact += float(actual_by_quarter.get(quarter, 0.0))
        rows.append(
            {
                "Период": f"Q{quarter}",
                "План накопительно": cumulative_plan,
                "Факт накопительно": cumulative_fact,
            }
        )
    return pd.DataFrame(rows)


def weekly_changes(connection: sqlite3.Connection, import_id: int) -> tuple[pd.DataFrame, dict[str, Any]]:
    current_meta = connection.execute(
        "SELECT * FROM imports WHERE id = ?", (import_id,)
    ).fetchone()
    previous = connection.execute(
        "SELECT id FROM imports WHERE id <> ? AND status = 'completed' "
        "AND COALESCE(week_end, imported_at) < COALESCE(?, ?) "
        "ORDER BY COALESCE(week_end, imported_at) DESC, id DESC LIMIT 1",
        (import_id, current_meta["week_end"], current_meta["imported_at"]),
    ).fetchone()
    if not previous:
        return pd.DataFrame(), {"previous_import_id": None}
    previous_id = int(previous["id"])
    current = deduplicate_projects(load_projects(connection, import_id)).set_index("business_key")
    old = deduplicate_projects(load_projects(connection, previous_id)).set_index("business_key")
    current_values = load_period_values(connection, import_id)
    old_values = load_period_values(connection, previous_id)
    current_actual = actual_by_project(current_values, int(str(current_meta["week_end"] or date.today())[:4]))
    old_meta = connection.execute("SELECT * FROM imports WHERE id = ?", (previous_id,)).fetchone()
    old_actual = actual_by_project(old_values, int(str(old_meta["week_end"] or date.today())[:4]))

    changes: list[dict[str, Any]] = []
    all_keys = set(current.index) | set(old.index)
    for key in all_keys:
        if key not in old.index:
            row = current.loc[key]
            changes.append({"Тип": "Новый проект", "Проект": row["project_name"], "Менеджер": row["manager"], "Изменение": "Появился впервые"})
            continue
        if key not in current.index:
            row = old.loc[key]
            changes.append({"Тип": "Исчез из отчёта", "Проект": row["project_name"], "Менеджер": row["manager"], "Изменение": "Нет в текущей версии"})
            continue
        now = current.loc[key]
        before = old.loc[key]
        if now["stage"] != before["stage"]:
            changes.append(
                {
                    "Тип": "Изменение стадии",
                    "Проект": now["project_name"],
                    "Менеджер": now["manager"],
                    "Изменение": f"{before['stage'] or 'не указана'} → {now['stage'] or 'не указана'}",
                }
            )
        now_actual = float(current_actual.get(key, 0.0))
        before_actual = float(old_actual.get(key, 0.0))
        if now_actual > before_actual + 0.01:
            changes.append(
                {
                    "Тип": "Отгрузка",
                    "Проект": now["project_name"],
                    "Менеджер": now["manager"],
                    "Изменение": now_actual - before_actual,
                }
            )
        if (now["planned_year"], now["planned_quarter"]) != (
            before["planned_year"], before["planned_quarter"]
        ):
            changes.append(
                {
                    "Тип": "Перенос",
                    "Проект": now["project_name"],
                    "Менеджер": now["manager"],
                    "Изменение": f"{before['planned_year']} Q{before['planned_quarter']} → {now['planned_year']} Q{now['planned_quarter']}",
                }
            )
    return pd.DataFrame(changes), {"previous_import_id": previous_id}


def forecast_history(connection: sqlite3.Connection, year: int) -> pd.DataFrame:
    imports = pd.read_sql_query(
        "SELECT id, filename, week_end, imported_at FROM imports WHERE status = 'completed' "
        "ORDER BY COALESCE(week_end, imported_at), id",
        connection,
    )
    rows: list[dict[str, Any]] = []
    for item in imports.to_dict("records"):
        kpis = calculate_kpis(connection, int(item["id"]), year)
        rows.append(
            {
                "Дата версии": item["week_end"] or item["imported_at"][:10],
                "Консервативный": kpis["conservative"],
                "Базовый": kpis["base_forecast"],
                "Оптимистичный": kpis["optimistic"],
                "Портфель": kpis["confirmed"] + kpis["signed"] + kpis["planned_contracts"],
            }
        )
    return pd.DataFrame(rows)


def quality_issues(connection: sqlite3.Connection, import_id: int) -> pd.DataFrame:
    projects = load_projects(connection, import_id)
    if projects.empty:
        return pd.DataFrame()
    issues: list[dict[str, Any]] = []

    def add(row: pd.Series, issue: str, severity: str) -> None:
        issues.append(
            {
                "Важность": severity,
                "Проблема": issue,
                "Проект": row.get("project_name"),
                "Менеджер": row.get("manager"),
                "Источник": f"{row.get('source_sheet')}!{int(row.get('source_row'))}",
                "Ключ": row.get("business_key"),
            }
        )

    for _, row in projects.iterrows():
        if pd.isna(row.get("project_amount")) or float(row.get("project_amount") or 0) == 0:
            add(row, "Отсутствует сумма проекта", "Ошибка")
        if not str(row.get("manager") or "").strip():
            add(row, "Не указан менеджер", "Ошибка")
        if row.get("stage") not in {"отгружено", "проиграно", "отменено"} and not row.get("planned_ship_date"):
            add(row, "Не указана точная дата отгрузки", "Предупреждение")
        if row.get("stage") == "отгружено" and not row.get("actual_ship_date"):
            add(row, "Проект отмечен как отгруженный, но нет фактической даты", "Ошибка")
        planned = pd.to_datetime(row.get("planned_ship_date"), errors="coerce")
        if pd.notna(planned) and planned.date() < date.today() and row.get("stage") != "отгружено":
            add(row, "Срок отгрузки уже прошёл", "Ошибка")
        if int(row.get("is_ambiguous") or 0):
            add(row, "Неоднозначное совпадение или несколько стадий", "Ошибка")

    values = load_period_values(connection, import_id)
    outstanding = current_outstanding(values, date.today().year)
    if not outstanding.empty:
        inactive = projects[projects["stage"].isin(INACTIVE_STAGES)][["business_key", "project_name", "manager", "source_sheet", "source_row"]]
        invalid = outstanding.merge(inactive, on="business_key")
        for _, row in invalid.iterrows():
            issues.append(
                {
                    "Важность": "Ошибка",
                    "Проблема": "Проигранный или отменённый проект попал в план отгрузок",
                    "Проект": row["project_name"],
                    "Менеджер": row["manager"],
                    "Источник": f"{row['source_sheet']}!{int(row['source_row'])}",
                    "Ключ": row["business_key"],
                }
            )
    return pd.DataFrame(issues).drop_duplicates() if issues else pd.DataFrame()


def export_excel(
    connection: sqlite3.Connection, import_id: int, year: int
) -> bytes:
    projects = deduplicate_projects(load_projects(connection, import_id))
    managers = manager_summary(connection, import_id, year)
    issues = quality_issues(connection, import_id)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        projects.to_excel(writer, sheet_name="Реестр проектов", index=False)
        managers.to_excel(writer, sheet_name="Менеджеры", index=False)
        issues.to_excel(writer, sheet_name="Контроль качества", index=False)
    return output.getvalue()


def executive_report(connection: sqlite3.Connection, import_id: int, year: int) -> str:
    kpi = calculate_kpis(connection, import_id, year)

    def millions(value: float) -> str:
        return f"{value / 1_000_000:,.2f}".replace(",", " ")

    return "\n".join(
        [
            f"Краткий отчёт руководителю за {year} год",
            "",
            f"Годовой план: {millions(kpi['plan'])} млн руб. без НДС.",
            f"Фактическая отгрузка: {millions(kpi['fact'])} млн руб. ({kpi['completion']:.1%} плана).",
            f"Подтверждённые отгрузки: {millions(kpi['confirmed'])} млн руб.",
            f"Базовый прогноз: {millions(kpi['base_forecast'])} млн руб.",
            f"Оптимистичный сценарий: {millions(kpi['optimistic'])} млн руб.",
            f"Подписанные, но не отгруженные контракты: {millions(kpi['signed'])} млн руб.",
            f"Потенциальный портфель следующего года: {millions(kpi['next_year'])} млн руб.",
            "",
            "Неизвестные вероятности не включены в риск-взвешенную часть базового прогноза.",
        ]
    )
