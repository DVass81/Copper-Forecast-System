import json
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook


APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
STATE_PATH = DATA_DIR / "copper_forecast_state.json"
OVERRIDES_PATH = DATA_DIR / "copper_forecast_overrides.json"
SETTINGS_PATH = DATA_DIR / "copper_forecast_settings.json"
LOCAL_WORKBOOK_PATH = APP_DIR / "March Copper Review Final.xlsx"
ONEDRIVE_WORKBOOK_PATH = (
    Path.home()
    / "OneDrive - Industrial Commutator Corporation, Inc"
    / "Desktop"
    / "March Copper Review Final.xlsx"
)

ACTIVE_MILL = "Tecnofil"
FUTURE_MILL = "Coppr Rod"
DIST_SOURCES = ("Williams", "Maverick")
DEFAULT_SOURCE_LEAD_WEEKS = {
    "Tecnofil": 11.0,
    "Coppr Rod": 9.0,
    "Williams": 13.0,
    "Maverick": 13.0,
}
DEFAULT_SAFETY_WEEKS_BY_MOVER = {
    "Fast": 14.0,
    "Medium": 10.0,
    "Slow": 6.0,
}


st.set_page_config(
    page_title="Copper Forecast System",
    page_icon="CF",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def inject_styles() -> None:
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 1.1rem;
                padding-bottom: 2rem;
                max-width: 1450px;
            }
            .main {
                background:
                    radial-gradient(circle at top right, rgba(212, 171, 70, 0.14), transparent 28%),
                    linear-gradient(180deg, #f5f8fc 0%, #eef3f8 100%);
            }
            .hero-card {
                background: linear-gradient(135deg, #123c6b 0%, #0f2743 100%);
                border: 1px solid rgba(214, 171, 70, 0.35);
                border-radius: 24px;
                padding: 1.35rem 1.45rem;
                margin-bottom: 1rem;
                box-shadow: 0 18px 40px rgba(8, 26, 44, 0.18);
            }
            .hero-card h1 {
                margin: 0 0 0.35rem 0;
                color: #fff7e2;
                font-size: 2rem;
            }
            .hero-card p {
                margin: 0;
                color: #d9e4ef;
                max-width: 820px;
                font-size: 1rem;
            }
            .mini-card {
                background: rgba(255, 255, 255, 0.94);
                border: 1px solid #d7e0e7;
                border-top: 4px solid #d6ab46;
                border-radius: 18px;
                padding: 0.9rem 1rem;
                box-shadow: 0 10px 24px rgba(18, 54, 75, 0.08);
            }
            .mini-card strong {
                color: #123c6b;
                font-size: 1.4rem;
            }
            .section-card {
                background: rgba(255, 255, 255, 0.96);
                border: 1px solid #d7e0e7;
                border-radius: 18px;
                padding: 1rem 1.1rem;
                box-shadow: 0 10px 24px rgba(18, 54, 75, 0.08);
                margin-bottom: 1rem;
            }
            .kpi-card {
                background: linear-gradient(180deg, rgba(18, 60, 107, 0.98) 0%, rgba(15, 39, 67, 0.98) 100%);
                border: 1px solid rgba(214, 171, 70, 0.38);
                border-radius: 18px;
                padding: 1rem;
                box-shadow: 0 10px 24px rgba(8, 26, 44, 0.14);
                margin-bottom: 0.85rem;
            }
            .kpi-value {
                color: #f0c75e;
                font-size: 1.65rem;
                font-weight: 800;
                line-height: 1.1;
            }
            .kpi-label {
                color: #dbe6f0;
                font-size: 0.9rem;
                margin-top: 0.25rem;
            }
            .item-hero {
                background: linear-gradient(135deg, #123c6b 0%, #0f2743 100%);
                border: 1px solid rgba(214, 171, 70, 0.35);
                border-radius: 18px;
                padding: 1rem 1.15rem;
                margin: 0 0 1rem 0;
                box-shadow: 0 12px 28px rgba(8, 26, 44, 0.18);
            }
            .item-hero h3 {
                margin: 0 0 0.25rem 0;
                color: #f5efe1;
            }
            .item-hero p {
                margin: 0;
                color: #d9e4ef;
            }
            .stButton > button,
            .stDownloadButton > button {
                border-radius: 12px;
                border: 1px solid #d2dbe4;
                box-shadow: 0 8px 20px rgba(18, 54, 75, 0.08);
            }
            div[data-baseweb="tab"] {
                border-radius: 12px 12px 0 0;
                background: rgba(255,255,255,0.75);
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def build_default_settings() -> dict:
    return {
        "source_lead_weeks": dict(DEFAULT_SOURCE_LEAD_WEEKS),
        "safety_weeks_by_mover": dict(DEFAULT_SAFETY_WEEKS_BY_MOVER),
        "preferred_mill": ACTIVE_MILL,
        "future_mill": FUTURE_MILL,
    }


def load_settings() -> dict:
    saved = load_json(SETTINGS_PATH)
    defaults = build_default_settings()
    defaults["source_lead_weeks"].update(saved.get("source_lead_weeks", {}))
    defaults["safety_weeks_by_mover"].update(saved.get("safety_weeks_by_mover", {}))
    defaults["preferred_mill"] = saved.get("preferred_mill", ACTIVE_MILL) or ACTIVE_MILL
    defaults["future_mill"] = saved.get("future_mill", FUTURE_MILL) or FUTURE_MILL
    return defaults


def resolve_default_workbook_path() -> Path:
    if LOCAL_WORKBOOK_PATH.exists():
        return LOCAL_WORKBOOK_PATH
    return ONEDRIVE_WORKBOOK_PATH


def clean_size(value: object) -> str:
    return str(value or "").strip()


def number(value: object) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_active_sizes(workbook) -> list[str]:
    if "Copper" not in workbook.sheetnames:
        return []
    worksheet = workbook["Copper"]
    sizes = []
    for row in worksheet.iter_rows(min_row=1, values_only=True):
        size = clean_size(row[0] if row else "")
        if size:
            sizes.append(size)
    return sizes


def parse_historical_usage(workbook) -> dict:
    if "Historical Usage" not in workbook.sheetnames:
        return {}
    worksheet = workbook["Historical Usage"]
    usage_map = {}
    for row in worksheet.iter_rows(min_row=5, max_col=11, values_only=True):
        size = clean_size(row[0] if row else "")
        if not size:
            continue
        annual_values = [number(value) for value in row[1:10] if isinstance(value, (int, float))]
        usage_map[size] = {
            "historical_total_lbs": sum(annual_values),
            "historical_avg_annual_lbs": (sum(annual_values) / len(annual_values)) if annual_values else 0.0,
        }
    return usage_map


def parse_copper_order(workbook, active_sizes: list[str]) -> list[dict]:
    worksheet = workbook["Copper Order"]
    active_set = set(active_sizes)
    items = []
    for row in worksheet.iter_rows(min_row=4, max_col=25, values_only=True):
        size = clean_size(row[0] if row else "")
        if not size or size.lower().startswith("total"):
            continue
        if active_set and size not in active_set:
            continue
        items.append(
            {
                "size": size,
                "inventory_7mo_ago_lbs": number(row[1]),
                "received_prev_7mo_lbs": number(row[2]),
                "usage_prev_7mo_lbs": number(row[3]),
                "alt_for_other_lbs": number(row[4]),
                "alt_used_for_size_lbs": number(row[5]),
                "adjusted_usage_7mo_lbs": number(row[6]) or number(row[3]),
                "usage_share": number(row[7]),
                "avg_monthly_lbs": number(row[8]),
                "williams_on_hand_lbs": number(row[9]),
                "williams_on_order_lbs": number(row[10]),
                "maverick_on_hand_lbs": number(row[11]),
                "maverick_on_order_lbs": number(row[12]),
                "sam_dong_on_order_lbs": number(row[13]),
                "tecnofil_on_order_lbs": number(row[14]),
                "total_on_order_lbs": number(row[15]),
                "williams_general_inventory_lbs": number(row[16]),
                "icc_inventory_current_lbs": number(row[17]),
                "jobs_pending_lbs": number(row[18]),
                "total_available_reported_lbs": number(row[19]),
                "avg_monthly_supply_lbs": number(row[20]),
                "large_jobs_prev_7mo_lbs": number(row[21]),
                "avg_monthly_supply_ex_large_lbs": number(row[22]),
            }
        )
    return items


def parse_workbook(workbook_bytes: bytes, source_name: str) -> dict:
    workbook = load_workbook(filename=BytesIO(workbook_bytes), data_only=True, read_only=True)
    active_sizes = parse_active_sizes(workbook)
    items = parse_copper_order(workbook, active_sizes)
    historical_usage = parse_historical_usage(workbook)
    for item in items:
        history = historical_usage.get(item["size"], {})
        item["historical_total_lbs"] = history.get("historical_total_lbs", 0.0)
        item["historical_avg_annual_lbs"] = history.get("historical_avg_annual_lbs", 0.0)
    return {
        "source_name": source_name,
        "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "items": items,
    }


def classify_movers(items: list[dict]) -> dict:
    ranked = sorted(
        [(item["size"], item.get("avg_monthly_lbs", 0.0) or 0.0) for item in items],
        key=lambda pair: pair[1],
    )
    count = len(ranked)
    mover_map = {}
    if not count:
        return mover_map
    for index, (size, _) in enumerate(ranked, start=1):
        position = index / count
        if position > 0.67:
            mover_map[size] = "Fast"
        elif position > 0.34:
            mover_map[size] = "Medium"
        else:
            mover_map[size] = "Slow"
    return mover_map


def choose_action(
    *,
    weekly_usage: float,
    plant_available_lbs: float,
    net_supply_lbs: float,
    dc_on_hand_lbs: float,
    reorder_point_lbs: float,
    target_stock_lbs: float,
    recommended_order_lbs: float,
) -> tuple[str, str]:
    if weekly_usage <= 0:
        if net_supply_lbs > 0:
            return "Excess Risk", "No current demand signal is showing, so this size should be reviewed for slow movement."
        return "Healthy", "No recent demand signal and no active supply issue."
    if plant_available_lbs < weekly_usage * 4 and dc_on_hand_lbs > 0:
        return "Pull From DC", "Plant coverage is short-term only, so distribution inventory should protect the schedule."
    if net_supply_lbs < reorder_point_lbs:
        return "Order Now", "Net supply will not cover mill lead time plus safety stock."
    if net_supply_lbs > target_stock_lbs * 1.6:
        return "Excess Risk", "Net supply is well above the current target range for this size."
    if recommended_order_lbs > 0:
        return "Monitor", "The item is above reorder point now, but it should stay on the buyer watch list."
    return "Healthy", "Supply position is currently inside the target range."


def planning_sort_key(row: dict) -> tuple:
    priority = {
        "Order Now": 0,
        "Pull From DC": 1,
        "Excess Risk": 2,
        "Monitor": 3,
        "Healthy": 4,
    }
    return (priority.get(row["action_bucket"], 9), -row["recommended_order_lbs"], row["size"])


def largest_dc_source(item: dict) -> str:
    if item.get("williams_on_hand_lbs", 0.0) >= item.get("maverick_on_hand_lbs", 0.0):
        return "Williams"
    return "Maverick"


def choose_source(source_override: str, recommendation_qty: float, preferred_mill: str) -> str:
    if source_override and source_override != "Auto":
        return source_override
    if recommendation_qty <= 0:
        return "None"
    return preferred_mill


def build_plan(snapshot: dict, overrides: dict, settings: dict) -> list[dict]:
    items = snapshot.get("items", [])
    lead_weeks = settings.get("source_lead_weeks", {})
    safety_by_mover = settings.get("safety_weeks_by_mover", {})
    preferred_mill = settings.get("preferred_mill", ACTIVE_MILL)
    future_mill = settings.get("future_mill", FUTURE_MILL)
    mover_map = classify_movers(items)
    planning_rows = []
    for item in items:
        override = overrides.get(item["size"], {})
        base_monthly_lbs = item.get("avg_monthly_lbs") or round(item.get("adjusted_usage_7mo_lbs", 0.0) / 7.0, 2)
        manual_adjustment_lbs = float(override.get("manual_adjustment_lbs", 0.0) or 0.0)
        forecast_monthly_lbs = max(0.0, base_monthly_lbs + manual_adjustment_lbs)
        weekly_usage = forecast_monthly_lbs / 4.345 if forecast_monthly_lbs else 0.0
        mover_class = mover_map.get(item["size"], "Slow")
        safety_weeks = float(safety_by_mover.get(mover_class, DEFAULT_SAFETY_WEEKS_BY_MOVER[mover_class]))

        plant_available_lbs = max(0.0, item.get("icc_inventory_current_lbs", 0.0) - item.get("jobs_pending_lbs", 0.0))
        dc_on_hand_lbs = item.get("williams_on_hand_lbs", 0.0) + item.get("maverick_on_hand_lbs", 0.0)
        dc_on_order_lbs = item.get("williams_on_order_lbs", 0.0) + item.get("maverick_on_order_lbs", 0.0)
        mill_on_order_lbs = item.get("tecnofil_on_order_lbs", 0.0)
        inbound_total_lbs = dc_on_order_lbs + mill_on_order_lbs
        current_supply_lbs = plant_available_lbs + dc_on_hand_lbs
        net_supply_lbs = current_supply_lbs + inbound_total_lbs

        safety_stock_lbs = weekly_usage * safety_weeks
        reorder_point_lbs = weekly_usage * float(lead_weeks.get(preferred_mill, 11.0)) + safety_stock_lbs
        target_stock_lbs = weekly_usage * (float(lead_weeks.get(preferred_mill, 11.0)) + safety_weeks + 4.0)
        current_coverage_weeks = (current_supply_lbs / weekly_usage) if weekly_usage else 999.0
        net_coverage_weeks = (net_supply_lbs / weekly_usage) if weekly_usage else 999.0
        recommendation_qty = max(0.0, target_stock_lbs - net_supply_lbs)
        action_bucket, recommendation_reason = choose_action(
            weekly_usage=weekly_usage,
            plant_available_lbs=plant_available_lbs,
            net_supply_lbs=net_supply_lbs,
            dc_on_hand_lbs=dc_on_hand_lbs,
            reorder_point_lbs=reorder_point_lbs,
            target_stock_lbs=target_stock_lbs,
            recommended_order_lbs=recommendation_qty,
        )
        source_override = override.get("source_override", "Auto")
        recommended_source = choose_source(source_override, recommendation_qty, preferred_mill)
        if action_bucket == "Pull From DC":
            recommendation_qty = min(dc_on_hand_lbs, max(0.0, (weekly_usage * 6.0) - plant_available_lbs))
            recommended_source = source_override if source_override not in {"", "Auto"} else largest_dc_source(item)
        elif action_bucket in {"Healthy", "Monitor", "Excess Risk"}:
            recommendation_qty = 0.0 if action_bucket != "Monitor" else recommendation_qty
            if action_bucket == "Excess Risk":
                recommended_source = "None"

        planning_rows.append(
            {
                "size": item["size"],
                "mover_class": mover_class,
                "base_monthly_lbs": base_monthly_lbs,
                "manual_adjustment_lbs": manual_adjustment_lbs,
                "forecast_monthly_lbs": forecast_monthly_lbs,
                "plant_available_lbs": plant_available_lbs,
                "dc_on_hand_lbs": dc_on_hand_lbs,
                "dc_on_order_lbs": dc_on_order_lbs,
                "mill_on_order_lbs": mill_on_order_lbs,
                "inbound_total_lbs": inbound_total_lbs,
                "net_supply_lbs": net_supply_lbs,
                "current_coverage_weeks": current_coverage_weeks,
                "net_coverage_weeks": net_coverage_weeks,
                "safety_stock_lbs": safety_stock_lbs,
                "reorder_point_lbs": reorder_point_lbs,
                "target_stock_lbs": target_stock_lbs,
                "recommended_order_lbs": recommendation_qty,
                "recommended_source": recommended_source,
                "action_bucket": action_bucket,
                "recommendation_reason": recommendation_reason,
                "jobs_pending_lbs": item.get("jobs_pending_lbs", 0.0),
                "override_note": override.get("note", "").strip(),
                "preferred_mill": preferred_mill,
                "future_mill": future_mill,
            }
        )
    return sorted(planning_rows, key=planning_sort_key)


def render_hero(snapshot: dict, plan_rows: list[dict]) -> None:
    st.markdown(
        """
        <div class="hero-card">
            <h1>Copper Forecast System</h1>
            <p>Track plant inventory, distribution stock, inbound copper, and reorder timing in pounds with direct-mill-first logic and planner overrides.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    shortage_count = len([row for row in plan_rows if row["action_bucket"] in {"Order Now", "Pull From DC"}])
    excess_count = len([row for row in plan_rows if row["action_bucket"] == "Excess Risk"])
    total_buy = sum(row["recommended_order_lbs"] for row in plan_rows)
    cards = [
        (len(snapshot.get("items", [])), "Active sizes"),
        (shortage_count, "Shortage actions"),
        (excess_count, "Excess risks"),
        (f"{total_buy:,.0f}", "Recommended lbs"),
    ]
    cols = st.columns(4)
    for col, (value, label) in zip(cols, cards):
        col.markdown(f'<div class="mini-card"><strong>{value}</strong><br>{label}</div>', unsafe_allow_html=True)


def render_dashboard(plan_rows: list[dict], snapshot: dict, settings: dict) -> None:
    if not plan_rows:
        st.info("Load your copper workbook in `Data Import / Refresh` to populate the planning dashboard.")
        return
    kpi_cols = st.columns(4)
    kpis = [
        ("Items in plan", len(plan_rows)),
        ("Order now", len([row for row in plan_rows if row["action_bucket"] == "Order Now"])),
        ("Pull from DC", len([row for row in plan_rows if row["action_bucket"] == "Pull From DC"])),
        ("Monitor", len([row for row in plan_rows if row["action_bucket"] == "Monitor"])),
    ]
    for col, (label, value) in zip(kpi_cols, kpis):
        col.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-value">{value}</div>
                <div class="kpi-label">{label}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    left, right = st.columns([1.3, 1.0])
    with left:
        st.markdown("#### Action Queue")
        rows = [
            {
                "Copper Size": row["size"],
                "Action": row["action_bucket"],
                "Source": row["recommended_source"],
                "Qty (lbs)": round(row["recommended_order_lbs"], 0),
                "Net Coverage (weeks)": round(row["net_coverage_weeks"], 1),
            }
            for row in plan_rows
            if row["action_bucket"] != "Healthy"
        ]
        st.dataframe(rows, use_container_width=True, hide_index=True)
    with right:
        st.markdown("#### Planning Rules")
        st.write(f"Preferred direct mill: {settings.get('preferred_mill', ACTIVE_MILL)}")
        st.write(f"Future mill: {settings.get('future_mill', FUTURE_MILL)}")
        st.write("Distribution centers are used as safety coverage when plant supply gets too tight.")
        st.write(f"Latest import: {snapshot.get('imported_at', 'Not loaded')}")
    chart_data = {row["size"]: row["recommended_order_lbs"] for row in plan_rows if row["recommended_order_lbs"] > 0}
    if chart_data:
        st.markdown("#### Top Recommended Buys")
        st.bar_chart(chart_data)


def render_items(plan_rows: list[dict]) -> None:
    if not plan_rows:
        st.info("No item details are available yet.")
        return
    row_map = {row["size"]: row for row in plan_rows}
    selected = st.selectbox("Copper size", list(row_map))
    row = row_map[selected]
    st.markdown(
        f"""
        <div class="item-hero">
            <h3>{row['size']}</h3>
            <p>{row['mover_class']} mover | {row['action_bucket']} | Preferred source: {row['recommended_source']}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    top = st.columns(4)
    top[0].metric("Forecast / month", f"{row['forecast_monthly_lbs']:,.0f} lbs")
    top[1].metric("Plant available", f"{row['plant_available_lbs']:,.0f} lbs")
    top[2].metric("DC on hand", f"{row['dc_on_hand_lbs']:,.0f} lbs")
    top[3].metric("Net coverage", f"{row['net_coverage_weeks']:.1f} wks")
    bottom = st.columns(4)
    bottom[0].metric("Safety stock", f"{row['safety_stock_lbs']:,.0f} lbs")
    bottom[1].metric("Reorder point", f"{row['reorder_point_lbs']:,.0f} lbs")
    bottom[2].metric("Target stock", f"{row['target_stock_lbs']:,.0f} lbs")
    bottom[3].metric("Recommended qty", f"{row['recommended_order_lbs']:,.0f} lbs")
    if row["override_note"]:
        st.info(f"Planner note: {row['override_note']}")


def render_supply_plan(plan_rows: list[dict]) -> None:
    if not plan_rows:
        st.info("No supply plan yet.")
        return
    rows = [
        {
            "Copper Size": row["size"],
            "Mover": row["mover_class"],
            "Plant Available": round(row["plant_available_lbs"], 0),
            "DC On Hand": round(row["dc_on_hand_lbs"], 0),
            "Inbound Total": round(row["inbound_total_lbs"], 0),
            "Forecast / Month": round(row["forecast_monthly_lbs"], 0),
            "Current Coverage": round(row["current_coverage_weeks"], 1),
            "Net Coverage": round(row["net_coverage_weeks"], 1),
            "Action": row["action_bucket"],
        }
        for row in plan_rows
    ]
    st.dataframe(rows, use_container_width=True, hide_index=True)


def render_recommendations(plan_rows: list[dict]) -> None:
    if not plan_rows:
        st.info("No recommendations yet.")
        return
    filter_value = st.selectbox(
        "Recommendation filter",
        ["All", "Order Now", "Pull From DC", "Excess Risk", "Monitor", "Healthy"],
    )
    rows = [
        {
            "Copper Size": row["size"],
            "Action": row["action_bucket"],
            "Source": row["recommended_source"],
            "Qty (lbs)": round(row["recommended_order_lbs"], 0),
            "Reorder Point": round(row["reorder_point_lbs"], 0),
            "Target Stock": round(row["target_stock_lbs"], 0),
            "Reason": row["recommendation_reason"],
        }
        for row in plan_rows
        if filter_value == "All" or row["action_bucket"] == filter_value
    ]
    st.dataframe(rows, use_container_width=True, hide_index=True)
    if rows:
        st.download_button(
            "Download recommendations as CSV",
            data=rows_to_csv(rows),
            file_name="copper_recommendations.csv",
            mime="text/csv",
        )


def render_overrides(plan_rows: list[dict], overrides: dict) -> None:
    if not plan_rows:
        st.info("Load data before editing overrides.")
        return
    selected_size = st.selectbox("Copper size to adjust", [row["size"] for row in plan_rows], key="override_size")
    existing = overrides.get(selected_size, {})
    with st.form("override_form"):
        adjustment = st.number_input(
            "Monthly demand adjustment (lbs)",
            value=float(existing.get("manual_adjustment_lbs", 0.0)),
            step=100.0,
        )
        source_override = st.selectbox(
            "Source override",
            ["Auto", ACTIVE_MILL, "Williams", "Maverick", FUTURE_MILL],
            index=["Auto", ACTIVE_MILL, "Williams", "Maverick", FUTURE_MILL].index(
                existing.get("source_override", "Auto")
            ),
        )
        note = st.text_area("Planner note", value=existing.get("note", ""), height=120)
        submitted = st.form_submit_button("Save override")
    if submitted:
        overrides[selected_size] = {
            "manual_adjustment_lbs": adjustment,
            "source_override": source_override,
            "note": note.strip(),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        save_json(OVERRIDES_PATH, overrides)
        st.success(f"Saved override for {selected_size}.")
        st.rerun()
    if overrides:
        st.dataframe(
            [
                {
                    "Copper Size": size,
                    "Monthly Adjustment": row.get("manual_adjustment_lbs", 0),
                    "Source Override": row.get("source_override", "Auto"),
                    "Note": row.get("note", ""),
                    "Updated": row.get("updated_at", ""),
                }
                for size, row in overrides.items()
            ],
            use_container_width=True,
            hide_index=True,
        )


def render_settings(settings: dict) -> None:
    st.markdown("#### Planning Assumptions")
    with st.form("settings_form"):
        preferred_mill = st.selectbox(
            "Preferred direct mill",
            [ACTIVE_MILL, FUTURE_MILL],
            index=[ACTIVE_MILL, FUTURE_MILL].index(settings.get("preferred_mill", ACTIVE_MILL)),
        )
        future_mill = st.selectbox(
            "Future / alternate mill",
            [FUTURE_MILL, ACTIVE_MILL],
            index=[FUTURE_MILL, ACTIVE_MILL].index(settings.get("future_mill", FUTURE_MILL)),
        )
        st.markdown("#### Lead Times (weeks)")
        lead_col1, lead_col2, lead_col3, lead_col4 = st.columns(4)
        tecnofil_weeks = lead_col1.number_input(
            "Tecnofil",
            value=float(settings["source_lead_weeks"].get("Tecnofil", 11.0)),
            step=0.5,
        )
        coppr_rod_weeks = lead_col2.number_input(
            "Coppr Rod",
            value=float(settings["source_lead_weeks"].get("Coppr Rod", 9.0)),
            step=0.5,
        )
        williams_weeks = lead_col3.number_input(
            "Williams",
            value=float(settings["source_lead_weeks"].get("Williams", 13.0)),
            step=0.5,
        )
        maverick_weeks = lead_col4.number_input(
            "Maverick",
            value=float(settings["source_lead_weeks"].get("Maverick", 13.0)),
            step=0.5,
        )
        st.markdown("#### Safety Stock by Mover (weeks)")
        safety_col1, safety_col2, safety_col3 = st.columns(3)
        fast_weeks = safety_col1.number_input(
            "Fast movers",
            value=float(settings["safety_weeks_by_mover"].get("Fast", 14.0)),
            step=1.0,
        )
        medium_weeks = safety_col2.number_input(
            "Medium movers",
            value=float(settings["safety_weeks_by_mover"].get("Medium", 10.0)),
            step=1.0,
        )
        slow_weeks = safety_col3.number_input(
            "Slow movers",
            value=float(settings["safety_weeks_by_mover"].get("Slow", 6.0)),
            step=1.0,
        )
        submitted = st.form_submit_button("Save planning settings")
    if submitted:
        save_json(
            SETTINGS_PATH,
            {
                "preferred_mill": preferred_mill,
                "future_mill": future_mill,
                "source_lead_weeks": {
                    "Tecnofil": tecnofil_weeks,
                    "Coppr Rod": coppr_rod_weeks,
                    "Williams": williams_weeks,
                    "Maverick": maverick_weeks,
                },
                "safety_weeks_by_mover": {
                    "Fast": fast_weeks,
                    "Medium": medium_weeks,
                    "Slow": slow_weeks,
                },
            },
        )
        st.success("Planning settings saved.")
        st.rerun()


def rows_to_csv(rows: list[dict]) -> str:
    if not rows:
        return ""
    headers = list(rows[0].keys())
    lines = [",".join(headers)]
    for row in rows:
        values = []
        for header in headers:
            value = str(row.get(header, ""))
            if "," in value or '"' in value or "\n" in value:
                value = '"' + value.replace('"', '""') + '"'
            values.append(value)
        lines.append(",".join(values))
    return "\n".join(lines)


def render_import(snapshot: dict) -> None:
    st.markdown("#### Workbook Import")
    default_workbook_path = resolve_default_workbook_path()
    uploaded_file = st.file_uploader("Upload copper workbook", type=["xlsx"])
    if uploaded_file and st.button("Import uploaded workbook"):
        parsed = parse_workbook(uploaded_file.getvalue(), uploaded_file.name)
        save_json(STATE_PATH, parsed)
        st.success(f"Imported {len(parsed.get('items', []))} copper sizes.")
        st.rerun()
    if default_workbook_path.exists() and st.button(f"Load default workbook: {default_workbook_path.name}"):
        parsed = parse_workbook(default_workbook_path.read_bytes(), default_workbook_path.name)
        save_json(STATE_PATH, parsed)
        st.success(f"Loaded {default_workbook_path.name}.")
        st.rerun()
    st.caption(str(default_workbook_path))
    if snapshot:
        st.dataframe(snapshot.get("items", [])[:10], use_container_width=True, hide_index=True)
    else:
        st.info("No workbook has been loaded yet.")


def main() -> None:
    inject_styles()
    snapshot = load_json(STATE_PATH)
    overrides = load_json(OVERRIDES_PATH)
    settings = load_settings()
    plan_rows = build_plan(snapshot, overrides, settings)
    render_hero(snapshot, plan_rows)
    dashboard_tab, items_tab, supply_tab, reorder_tab, overrides_tab, settings_tab, import_tab = st.tabs(
        [
            "Dashboard",
            "Copper Items",
            "Supply Plan",
            "Reorder Recommendations",
            "Manual Forecast Overrides",
            "Planning Settings",
            "Data Import / Refresh",
        ]
    )
    with dashboard_tab:
        render_dashboard(plan_rows, snapshot, settings)
    with items_tab:
        render_items(plan_rows)
    with supply_tab:
        render_supply_plan(plan_rows)
    with reorder_tab:
        render_recommendations(plan_rows)
    with overrides_tab:
        render_overrides(plan_rows, overrides)
    with settings_tab:
        render_settings(settings)
    with import_tab:
        render_import(snapshot)


if __name__ == "__main__":
    main()
