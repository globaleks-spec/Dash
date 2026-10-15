from __future__ import annotations

import os
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from dashboard.analytics import (
    apply_filters,
    calculate_kpis,
    current_outstanding,
    deduplicate_projects,
    executive_report,
    export_excel,
    forecast_history,
    load_period_values,
    load_projects,
    manager_summary,
    quality_issues,
    quarterly_cumulative,
    weekly_changes,
)
from dashboard.db import (
    DEFAULT_DB_PATH,
    connect,
    initialize,
    list_imports,
    save_override,
)
from dashboard.excel_import import MappingRequired, import_workbook, profile_workbook


st.set_page_config(
    page_title="Продажи · Завод Инвертор",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --ink: #17202a; --muted: #637083; --line: #dfe5eb; --accent: #0b5cab; }
    .stApp { background: #f5f7f9; color: var(--ink); }
    [data-testid="stHeader"] { background: rgba(245,247,249,.92); }
    [data-testid="stSidebar"] { background: #eef2f5; border-right: 1px solid var(--line); }
    div[data-testid="stMetric"] {
        background: white; border: 1px solid var(--line); border-radius: 8px;
        padding: 14px 16px; min-height: 108px; box-shadow: 0 1px 2px rgba(12,32,56,.05);
    }
    div[data-testid="stMetric"] label { color: var(--muted); font-size: .82rem; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: var(--ink); }
    .source-note { color: #536273; font-size: .86rem; margin-top: -8px; margin-bottom: 10px; }
    .section-note { background: #fff; border-left: 3px solid #0b5cab; padding: 10px 14px; color: #425166; }
    .status-good { color: #147a43; } .status-risk { color: #a36500; }
    .status-bad { color: #b42318; } .status-unknown { color: #667085; }
    h1, h2, h3 { color: #152536; letter-spacing: -.015em; }
    div[data-testid="stDataFrame"] { background: white; border: 1px solid var(--line); border-radius: 8px; }
    </style>
    """,
    unsafe_allow_html=True,
)


DB_PATH = Path(os.getenv("INVERTER_DB_PATH", str(DEFAULT_DB_PATH)))
DEMO_MODE = os.getenv("INVERTER_DEMO", "").strip().lower() in {"1", "true", "yes"}
connection = connect(DB_PATH)
initialize(connection)

if DEMO_MODE:
    st.warning(
        "ДЕМОНСТРАЦИОННАЯ ВЕРСИЯ · Все организации, проекты и суммы являются "
        "вымышленными. Загрузка реальных файлов отключена. Изменения могут быть "
        "сброшены при перезапуске облачного приложения."
    )


def rub_m(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{float(value) / 1_000_000:,.2f} млн ₽".replace(",", " ")


def percent(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{float(value):.1%}"


def safe_options(frame: pd.DataFrame, column: str) -> list[str]:
    if frame.empty or column not in frame:
        return []
    return sorted(frame[column].fillna("Не указано").astype(str).unique().tolist())


def common_filters(projects: pd.DataFrame) -> tuple[dict[str, Any], int | None, int | None]:
    st.caption("Фильтры применяются к показателям и таблицам текущего экрана.")
    row1 = st.columns([1.0, 1.0, 1.0, 1.4])
    quarter_label = row1[0].selectbox("Квартал", ["Все", "Q1", "Q2", "Q3", "Q4"])
    month_label = row1[1].selectbox(
        "Месяц",
        ["Все", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"],
    )
    manager = row1[2].multiselect("Менеджер", safe_options(projects, "manager"))
    customer = row1[3].multiselect("Заказчик", safe_options(projects, "customer"))
    row2 = st.columns([1.0, 1.0, 1.0, 1.4])
    industry = row2[0].multiselect("Отрасль", safe_options(projects, "industry"))
    channel = row2[1].multiselect("Канал", safe_options(projects, "channel"))
    equipment = row2[2].multiselect("Оборудование", safe_options(projects, "equipment_type"))
    search = row2[3].text_input("Поиск по проектам", placeholder="Заказчик, объект, номер счёта")
    filters = {
        "manager": manager,
        "customer": customer,
        "industry": industry,
        "channel": channel,
        "equipment_type": equipment,
        "search": search,
    }
    quarter = None if quarter_label == "Все" else int(quarter_label[1])
    month_names = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
    ]
    month = None if month_label == "Все" else month_names.index(month_label) + 1
    filters["month"] = month
    return filters, quarter, month


def kpi_grid(kpi: dict[str, float]) -> None:
    first = st.columns(5)
    first[0].metric("Годовой план", rub_m(kpi["plan"]))
    first[1].metric("Факт отгрузки", rub_m(kpi["fact"]), percent(kpi["completion"]))
    first[2].metric("Выполнение плана", percent(kpi["completion"]))
    first[3].metric("Остаток до плана", rub_m(kpi["gap"]))
    first[4].metric("Подтверждённые отгрузки", rub_m(kpi["confirmed"]))
    second = st.columns(4)
    second[0].metric("Прогноз с учётом рисков", rub_m(kpi["base_forecast"]))
    second[1].metric("Подписанные контракты", rub_m(kpi["signed"]))
    second[2].metric("План заключения договоров", rub_m(kpi["planned_contracts"]))
    second[3].metric("Портфель следующего года", rub_m(kpi["next_year"]))


def plan_chart(kpi: dict[str, float]) -> go.Figure:
    probable = max(kpi["base_forecast"] - kpi["conservative"], 0.0)
    gap = max(kpi["plan"] - kpi["base_forecast"], 0.0)
    overflow = max(kpi["base_forecast"] - kpi["plan"], 0.0)
    segments = [
        ("Факт", kpi["fact"], "#17864b"),
        ("Подтверждено", kpi["confirmed"], "#2f80ed"),
        ("Вероятно", probable, "#f2b84b"),
        ("Недостаёт", gap, "#d9dee5"),
    ]
    figure = go.Figure()
    for name, value, color in segments:
        figure.add_bar(
            y=["Годовой план"], x=[value / 1_000_000], name=name,
            orientation="h", marker_color=color,
            hovertemplate=f"{name}: %{{x:,.2f}} млн ₽<extra></extra>",
        )
    if overflow:
        figure.add_bar(
            y=["Сверх плана"], x=[overflow / 1_000_000], name="Сверх плана",
            orientation="h", marker_color="#7a5af8",
        )
    figure.update_layout(
        barmode="stack", height=250, margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", y=1.12, x=0),
        xaxis_title="млн руб. без НДС", yaxis_title=None,
        paper_bgcolor="white", plot_bgcolor="white",
    )
    return figure


def page_summary(import_id: int, year: int, projects: pd.DataFrame) -> None:
    st.title("Сводка руководителя")
    filters, quarter, _ = common_filters(projects)
    kpi = calculate_kpis(connection, import_id, year, filters, quarter)
    kpi_grid(kpi)
    left, right = st.columns([1.15, 1])
    with left:
        st.subheader("Выполнение годового плана")
        st.plotly_chart(plan_chart(kpi), width="stretch")
        st.caption(
            "Базовый сценарий учитывает факт, подтверждённые отгрузки и только явно заданные вероятности. "
            "Неизвестная вероятность не заменяется автоматическим значением."
        )
    with right:
        st.subheader("План и факт накопительно")
        cumulative = quarterly_cumulative(connection, import_id, year, filters)
        if not cumulative.empty:
            melted = cumulative.melt("Период", var_name="Показатель", value_name="Сумма")
            fig = px.line(
                melted, x="Период", y=melted["Сумма"] / 1_000_000,
                color="Показатель", markers=True,
                color_discrete_map={"План накопительно": "#0b5cab", "Факт накопительно": "#17864b"},
            )
            fig.update_layout(
                height=310, margin=dict(l=20, r=20, t=20, b=20),
                yaxis_title="млн руб. без НДС", xaxis_title=None,
                paper_bgcolor="white", plot_bgcolor="white",
            )
            st.plotly_chart(fig, width="stretch")
            st.caption("В исходной книге план и факт разбиты по кварталам, поэтому месячная линия не интерполируется.")
    filtered = apply_filters(projects, filters)
    risk_columns = ["risk_financial", "risk_technical", "risk_production", "risk_procurement", "risk_schedule"]
    if not filtered.empty:
        risky = filtered[filtered[risk_columns].fillna(0).astype(int).sum(axis=1) > 0]
        if not risky.empty:
            st.subheader("Проекты с явно указанными рисками")
            st.dataframe(
                risky[["manager", "project_name", "stage", "project_amount", *risk_columns]].rename(
                    columns={
                        "manager": "Менеджер", "project_name": "Проект", "stage": "Стадия",
                        "project_amount": "Сумма", "risk_financial": "Финансы",
                        "risk_technical": "Техника", "risk_production": "Производство",
                        "risk_procurement": "Закупки", "risk_schedule": "Сроки",
                    }
                ),
                hide_index=True, width="stretch",
                column_config={"Сумма": st.column_config.NumberColumn(format="%.0f ₽")},
            )


def page_managers(import_id: int, year: int, projects: pd.DataFrame) -> None:
    st.title("Аналитика менеджеров")
    quarter_label = st.selectbox("План отгрузок за квартал", ["Текущий", "Q1", "Q2", "Q3", "Q4"])
    current_quarter = (date.today().month - 1) // 3 + 1
    quarter = current_quarter if quarter_label == "Текущий" else int(quarter_label[1])
    summary = manager_summary(connection, import_id, year, quarter)
    if summary.empty:
        st.info("Нет данных по менеджерам.")
        return
    st.dataframe(
        summary,
        hide_index=True,
        width="stretch",
        column_config={
            "% выполнения": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=max(1.0, float(summary["% выполнения"].max()))),
            **{
                column: st.column_config.NumberColumn(format="%.0f ₽")
                for column in ["Годовой план", "Факт отгрузки", "Подписанные контракты", "План отгрузок квартала", "Прогноз до конца года", "План заключения договоров", "Портфель следующего года"]
            },
        },
    )
    selected = st.selectbox("Персональная панель", summary["Менеджер"].tolist())
    manager_projects = projects[projects["manager"] == selected]
    kpi = calculate_kpis(connection, import_id, year, {"manager": [selected]})
    cols = st.columns(4)
    cols[0].metric("План", rub_m(kpi["plan"]))
    cols[1].metric("Факт", rub_m(kpi["fact"]), percent(kpi["completion"]))
    cols[2].metric("Базовый прогноз", rub_m(kpi["base_forecast"]))
    cols[3].metric("Следующий год", rub_m(kpi["next_year"]))
    st.dataframe(
        manager_projects[["business_key", "project_name", "stage", "project_amount", "planned_year", "planned_quarter", "probability"]].rename(
            columns={"business_key": "ID", "project_name": "Проект", "stage": "Стадия", "project_amount": "Сумма", "planned_year": "Год", "planned_quarter": "Квартал", "probability": "Вероятность"}
        ),
        hide_index=True, width="stretch",
        column_config={"Сумма": st.column_config.NumberColumn(format="%.0f ₽"), "Вероятность": st.column_config.NumberColumn(format="%.0f%%")},
    )


EDIT_FIELDS = [
    "manager", "customer", "final_customer", "project_name", "channel", "industry",
    "equipment_type", "stage", "contract_date", "planned_ship_date", "actual_ship_date",
    "probability", "next_owner", "next_action", "next_action_due", "comment",
]


def page_projects(import_id: int, projects: pd.DataFrame) -> None:
    st.title("Реестр проектов")
    filters, _, _ = common_filters(projects)
    filtered = apply_filters(projects, filters)
    st.caption(f"Проектов после фильтрации: {len(filtered):,}".replace(",", " "))
    display = filtered[["business_key", *EDIT_FIELDS, "project_amount", "planned_year", "planned_quarter", "source_sheet", "source_row", "is_ambiguous"]].copy()
    display = display.rename(
        columns={
            "business_key": "ID", "manager": "Менеджер", "customer": "Заказчик",
            "final_customer": "Конечный заказчик", "project_name": "Проект", "channel": "Канал",
            "industry": "Отрасль", "equipment_type": "Оборудование", "stage": "Стадия",
            "contract_date": "Дата договора", "planned_ship_date": "Плановая отгрузка",
            "actual_ship_date": "Фактическая отгрузка", "probability": "Вероятность",
            "next_owner": "Ответственный за действие", "next_action": "Следующее действие",
            "next_action_due": "Срок действия", "comment": "Комментарий", "project_amount": "Сумма",
            "planned_year": "Год", "planned_quarter": "Квартал", "source_sheet": "Лист",
            "source_row": "Строка", "is_ambiguous": "На проверке",
        }
    )
    edited = st.data_editor(
        display,
        hide_index=True,
        width="stretch",
        height=560,
        disabled=["ID", "Сумма", "Год", "Квартал", "Лист", "Строка", "На проверке"],
        column_config={
            "ID": st.column_config.TextColumn(width="small"),
            "Проект": st.column_config.TextColumn(width="large"),
            "Сумма": st.column_config.NumberColumn(format="%.0f ₽"),
            "Вероятность": st.column_config.NumberColumn(min_value=0.0, max_value=1.0, step=0.05, format="%.0f%%"),
            "Стадия": st.column_config.SelectboxColumn(
                options=["техническая проработка", "ТКП", "участие в тендере", "техническая оценка", "акцепт", "подписанный договор", "производство", "изготовлено", "отгружено", "проиграно", "отменено", "перенесено"]
            ),
            "Канал": st.column_config.SelectboxColumn(options=["Прямые продажи", "СЭЩ", "КИТ", "Агент", "Другой"]),
        },
        key="project_editor",
    )
    if st.button("Сохранить изменения", type="primary"):
        original = display.set_index("ID")
        changed = 0
        reverse = {display_name: field for field, display_name in {
            "manager": "Менеджер", "customer": "Заказчик", "final_customer": "Конечный заказчик",
            "project_name": "Проект", "channel": "Канал", "industry": "Отрасль",
            "equipment_type": "Оборудование", "stage": "Стадия", "contract_date": "Дата договора",
            "planned_ship_date": "Плановая отгрузка", "actual_ship_date": "Фактическая отгрузка",
            "probability": "Вероятность", "next_owner": "Ответственный за действие",
            "next_action": "Следующее действие", "next_action_due": "Срок действия", "comment": "Комментарий",
        }.items()}
        for _, row in edited.iterrows():
            key = row["ID"]
            updates: dict[str, Any] = {}
            for display_name, field in reverse.items():
                before = original.loc[key, display_name]
                after = row[display_name]
                before_blank = pd.isna(before) or before == ""
                after_blank = pd.isna(after) or after == ""
                if (before_blank and after_blank) or (not before_blank and not after_blank and before == after):
                    continue
                updates[field] = None if after_blank else after
            if updates:
                save_override(connection, key, updates)
                changed += 1
        st.success(f"Сохранено проектов: {changed}")
        st.rerun()


def page_sesh(import_id: int, year: int, projects: pd.DataFrame) -> None:
    st.title("Проекты через АО «СЭЩ»")
    mask = projects["channel"].eq("СЭЩ") | projects["source_text"].fillna("").str.contains("СЭЩ|Электрощит", case=False, regex=True)
    sesh = projects[mask].copy()
    if sesh.empty:
        st.info("В текущей версии проекты СЭЩ не обнаружены.")
        return
    customers = safe_options(sesh, "final_customer")
    selected = st.multiselect("Конечный заказчик", customers)
    if selected:
        sesh = sesh[sesh["final_customer"].fillna("Не указано").isin(selected)]
    keys = sesh["business_key"].tolist()
    values = load_period_values(connection, import_id, keys)
    outstanding = current_outstanding(values, year)
    actual = values[(values["year"] == year) & (values["metric"] == "actual")]
    total_portfolio = float(sesh["project_amount"].fillna(0).sum())
    signed = float(outstanding.loc[outstanding["metric"] == "signed", "amount"].sum()) if not outstanding.empty else 0.0
    shipped = float(actual.groupby(["business_key", "quarter"])["amount"].max().sum()) if not actual.empty else 0.0
    next_year = float(current_outstanding(values, year + 1)["amount"].sum())
    cols = st.columns(4)
    cols[0].metric("Общий портфель", rub_m(total_portfolio))
    cols[1].metric("Подписанные договоры", rub_m(signed))
    cols[2].metric("Факт отгрузки", rub_m(shipped))
    cols[3].metric("Перспектива следующего года", rub_m(next_year))
    cols2 = st.columns(4)
    cols2[0].metric("В производстве", int(sesh["stage"].isin(["производство", "изготовлено"]).sum()))
    cols2[1].metric("Активные тендеры", int(sesh["stage"].eq("участие в тендере").sum()))
    cols2[2].metric("Проиграно", int(sesh["stage"].eq("проиграно").sum()))
    cols2[3].metric("Без актуального статуса", int(sesh["stage"].isna().sum()))
    left, right = st.columns(2)
    with left:
        chart = sesh.assign(**{"Сумма": sesh["project_amount"].fillna(0) / 1_000_000}).groupby("final_customer", dropna=False)["Сумма"].sum().sort_values(ascending=False).head(12).reset_index()
        chart["final_customer"] = chart["final_customer"].fillna("Не указан")
        st.plotly_chart(px.bar(chart, x="Сумма", y="final_customer", orientation="h", labels={"final_customer": "Конечный заказчик", "Сумма": "млн ₽"}), width="stretch")
    with right:
        chart = sesh.assign(**{"Сумма": sesh["project_amount"].fillna(0) / 1_000_000}).groupby("manager", dropna=False)["Сумма"].sum().sort_values(ascending=False).reset_index()
        st.plotly_chart(px.bar(chart, x="manager", y="Сумма", labels={"manager": "Менеджер", "Сумма": "млн ₽"}), width="stretch")
    st.dataframe(sesh[["manager", "final_customer", "project_name", "stage", "project_amount", "planned_year", "planned_quarter"]], hide_index=True, width="stretch")


def page_forecast(import_id: int, year: int, projects: pd.DataFrame) -> None:
    st.title("Прогноз и риски")
    kpi = calculate_kpis(connection, import_id, year)
    scenarios = pd.DataFrame(
        {
            "Сценарий": ["Консервативный", "Базовый", "Оптимистичный"],
            "Сумма": [kpi["conservative"], kpi["base_forecast"], kpi["optimistic"]],
        }
    )
    fig = px.bar(
        scenarios, x="Сценарий", y=scenarios["Сумма"] / 1_000_000,
        color="Сценарий", text_auto=".2f",
        color_discrete_sequence=["#2f80ed", "#f2b84b", "#17864b"],
        labels={"y": "млн руб. без НДС"},
    )
    fig.update_layout(showlegend=False, height=360, paper_bgcolor="white", plot_bgcolor="white")
    st.plotly_chart(fig, width="stretch")
    st.caption(f"Вероятность задана для {kpi['probability_known_share']:.1%} активных проектов портфеля.")

    editable = projects[~projects["stage"].isin(["отгружено", "проиграно", "отменено"])][["business_key", "manager", "project_name", "stage", "project_amount", "probability"]].copy()
    edited = st.data_editor(
        editable.rename(columns={"business_key": "ID", "manager": "Менеджер", "project_name": "Проект", "stage": "Стадия", "project_amount": "Сумма", "probability": "Вероятность"}),
        hide_index=True, width="stretch", height=360,
        disabled=["ID", "Менеджер", "Проект", "Стадия", "Сумма"],
        column_config={"Сумма": st.column_config.NumberColumn(format="%.0f ₽"), "Вероятность": st.column_config.NumberColumn(min_value=0.0, max_value=1.0, step=0.05, format="%.0f%%")},
        key="probability_editor",
    )
    if st.button("Сохранить вероятности", type="primary"):
        before = editable.set_index("business_key")["probability"]
        changed = 0
        for _, row in edited.iterrows():
            value = row["Вероятность"]
            old = before.get(row["ID"])
            if (pd.isna(value) and pd.isna(old)) or (pd.notna(value) and pd.notna(old) and float(value) == float(old)):
                continue
            save_override(connection, row["ID"], {"probability": None if pd.isna(value) else float(value)})
            changed += 1
        st.success(f"Вероятности сохранены: {changed}")
        st.rerun()

    st.subheader("Риск переноса IV квартала на следующий год")
    transfer = projects[
        ((projects["planned_year"] == year) & (projects["planned_quarter"] == 4) & projects["risk_schedule"].fillna(0).astype(bool))
        | (projects["planned_year"] > year)
        | projects["stage"].eq("перенесено")
    ]
    st.dataframe(transfer[["manager", "project_name", "stage", "project_amount", "planned_year", "planned_quarter", "source_text"]], hide_index=True, width="stretch")


def page_history(import_id: int, year: int) -> None:
    st.title("Динамика по неделям")
    history = forecast_history(connection, year)
    if len(history) > 1:
        melted = history.melt("Дата версии", var_name="Показатель", value_name="Сумма")
        fig = px.line(melted, x="Дата версии", y=melted["Сумма"] / 1_000_000, color="Показатель", markers=True, labels={"y": "млн руб. без НДС"})
        fig.update_layout(height=380, paper_bgcolor="white", plot_bgcolor="white")
        st.plotly_chart(fig, width="stretch")
    else:
        st.info("Для графика нужна как минимум вторая недельная версия. Каждая новая загрузка сохраняется отдельно.")
    changes, meta = weekly_changes(connection, import_id)
    if meta.get("previous_import_id") is None:
        st.info("Предыдущая версия не найдена.")
    elif changes.empty:
        st.success("Изменений по идентифицированным проектам не обнаружено.")
    else:
        counts = changes["Тип"].value_counts()
        cols = st.columns(min(5, len(counts)))
        for index, (kind, count) in enumerate(counts.head(5).items()):
            cols[index].metric(kind, int(count))
        st.dataframe(changes, hide_index=True, width="stretch")


def page_quality(import_id: int) -> None:
    st.title("Контроль качества данных")
    issues = quality_issues(connection, import_id)
    if issues.empty:
        st.success("Проблемы по заданным правилам не обнаружены.")
        return
    counts = issues["Важность"].value_counts()
    cols = st.columns(3)
    cols[0].metric("Всего замечаний", len(issues))
    cols[1].metric("Ошибки", int(counts.get("Ошибка", 0)))
    cols[2].metric("Предупреждения", int(counts.get("Предупреждение", 0)))
    severity = st.multiselect("Важность", sorted(issues["Важность"].unique()), default=sorted(issues["Важность"].unique()))
    st.dataframe(issues[issues["Важность"].isin(severity)], hide_index=True, width="stretch", height=600)


def page_import(imports: list[dict[str, Any]]) -> None:
    st.title("Загрузка и версии данных")
    if DEMO_MODE:
        st.info(
            "В демонстрационной версии загрузка Excel отключена, чтобы реальные "
            "коммерческие данные не могли попасть в тестовое облако."
        )
        if imports:
            st.dataframe(pd.DataFrame(imports), hide_index=True, width="stretch")
        return
    uploaded = st.file_uploader("Выберите еженедельный Excel-отчёт", type=["xlsx", "xlsm"])
    if uploaded:
        data = uploaded.getvalue()
        try:
            profile = profile_workbook(data)
        except Exception as error:
            st.error(f"Не удалось прочитать книгу: {error}")
            return
        st.subheader("Обнаруженная структура")
        sheet_frame = pd.DataFrame(
            [{"Лист": item["name"], "Состояние": item["state"], "Строк": item["rows"], "Столбцов": item["columns"], "Строка заголовка": item["header_candidates"][0]["row"] if item["header_candidates"] else None} for item in profile["sheets"]]
        )
        st.dataframe(sheet_frame, hide_index=True, width="stretch")
        st.write(f"Основной лист: **{profile['primary_sheet'] or 'не определён'}**")
        if profile["suggested_mapping"]:
            st.subheader("Предлагаемое соответствие полей")
            st.dataframe(pd.DataFrame(profile["suggested_mapping"]), hide_index=True, width="stretch")
        if st.button("Загрузить как новую недельную версию", type="primary"):
            try:
                result = import_workbook(connection, data, uploaded.name)
            except MappingRequired as error:
                st.error(str(error))
            except Exception as error:
                st.exception(error)
            else:
                if result.already_imported:
                    st.warning(f"Этот файл уже загружен. Версия №{result.import_id} не продублирована.")
                else:
                    st.success(f"Версия №{result.import_id} загружена: {result.projects} проектов, {result.managers} менеджеров.")
                st.rerun()
    st.subheader("История загрузок")
    if imports:
        st.dataframe(pd.DataFrame(imports), hide_index=True, width="stretch")
    else:
        st.info("Пока нет загруженных версий.")


def page_export(import_id: int, year: int) -> None:
    st.title("Экспорт и краткий отчёт")
    st.write("Экспорт содержит текущий реестр проектов с ручными корректировками, аналитику менеджеров и контроль качества.")
    left, right = st.columns(2)
    with left:
        if st.button("Подготовить Excel"):
            st.session_state["excel_export"] = export_excel(connection, import_id, year)
        if "excel_export" in st.session_state:
            st.download_button(
                "Скачать Excel",
                st.session_state["excel_export"],
                file_name=f"Инвертор_дашборд_{year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
    with right:
        report = executive_report(connection, import_id, year)
        st.download_button(
            "Скачать краткий отчёт",
            report.encode("utf-8-sig"),
            file_name=f"Краткий_отчет_{year}.txt",
            mime="text/plain",
        )
    st.text_area("Предпросмотр", executive_report(connection, import_id, year), height=280)


imports = list_imports(connection)
st.sidebar.title("АО «Завод Инвертор»")
st.sidebar.caption("Управленческий дашборд продаж")
page = st.sidebar.radio(
    "Раздел",
    ["Сводка", "Менеджеры", "Проекты", "СЭЩ", "Прогноз и риски", "Динамика", "Качество данных", "Загрузка", "Экспорт"],
)

if not imports:
    page_import(imports)
    st.stop()

version_labels = {
    f"{item['week_end'] or item['imported_at'][:10]} · {item['filename']}": item["id"]
    for item in imports
}
selected_label = st.sidebar.selectbox("Версия данных", list(version_labels))
import_id = int(version_labels[selected_label])
selected_meta = next(item for item in imports if item["id"] == import_id)

years = [
    row[0]
    for row in connection.execute(
        "SELECT DISTINCT year FROM period_values WHERE import_id = ? ORDER BY year", (import_id,)
    ).fetchall()
]
default_year = date.today().year if date.today().year in years else (years[0] if years else date.today().year)
year = st.sidebar.selectbox("Год", years or [default_year], index=(years.index(default_year) if default_year in years else 0))
st.sidebar.caption(f"Источник: {selected_meta['source_unit'] or 'единица не определена'}")

projects = deduplicate_projects(load_projects(connection, import_id))

if page == "Сводка":
    page_summary(import_id, year, projects)
elif page == "Менеджеры":
    page_managers(import_id, year, projects)
elif page == "Проекты":
    page_projects(import_id, projects)
elif page == "СЭЩ":
    page_sesh(import_id, year, projects)
elif page == "Прогноз и риски":
    page_forecast(import_id, year, projects)
elif page == "Динамика":
    page_history(import_id, year)
elif page == "Качество данных":
    page_quality(import_id)
elif page == "Загрузка":
    page_import(imports)
else:
    page_export(import_id, year)
