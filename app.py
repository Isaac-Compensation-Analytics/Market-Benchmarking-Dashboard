
"""
Market Benchmarking Dashboard (Dash) — LOCAL MODE (single Excel file, 2 tabs)

UPDATED RULES
- ALL roles are treated as SALARY
    Base Pay = Base Pay (no annualizing)
    Market Base Pay = Market Base Pay (already annual)
- Bonus Paid + Total Compensation:
    Used AS-IS (no annualizing)
- Hourly logic removed entirely

Updates included:
- Top filters: Metric, Job Title, Area Differential (Area is far right)
- Year-End Performance Rating filter applies ONLY to Costing Scenarios when Metric = Base Pay
  (logic unchanged; filter is placed inside the Costing Scenarios section visually)
- Employee − Market Delta table:
    - Headcount removed from the table
    - "Total Employee Headcount = X" shown under the table
- Market = National by default + Area Differential adjustment:
    High +12%, Mid 0%, Low -5%
- Injects realistic employee variability tied to Area Differential + Year-End Performance Rating

Breakouts (restructured):
- Breakouts are NOT linked to top filters
- Breakouts are driven by:
    - Breakout Metric
    - Sort by Δ Percentile
    - Sort Order
    - Job Family filter
- Output: ALL Job Titles within the selected Job Family, with Headcount and ΔP10..ΔP90
  (no max job cap)
"""

import os
import re
import numpy as np
import pandas as pd

from dash import Dash, dcc, html, Input, Output
import plotly.graph_objects as go
import plotly.express as px


# -----------------------------
# FILE (single workbook with 2 tabs)
# -----------------------------
EMPLOYEE_FILE = os.getenv("EMPLOYEE_FILE", "capital_one_employee_population.xlsx")
EMPLOYEE_SHEET = "Employee Source Data"
MARKET_SHEET = "National Market Data"

# Percentiles
PCTS = ["P10", "P25", "P50", "P75", "P90"]
PCT_Q = {"P10": 0.10, "P25": 0.25, "P50": 0.50, "P75": 0.75, "P90": 0.90}

REQ_EMP_COLS = [
    "Employee ID",
    "Job Family",
    "Job Title",
    "Job Level",
    "Area Differential",
    "Base Pay",
    "Bonus Paid",
    "Total Compensation",
    "Performance Rating",
]

# -----------------------------
# Theme
# -----------------------------
BRAND_RED = "#B00020"
BRAND_DARK = "#1F2937"
BRAND_GRAY = "#6B7280"
BRAND_BG = "#F6F7F9"
CARD_BG = "#FFFFFF"
BORDER = "#E5E7EB"
PLOT_TEMPLATE = "plotly_white"

# Market Area Differential factors
AREA_FACTORS = {"High": 1.12, "Mid": 1.00, "Low": 0.95}

# Performance mapping (Year-End Performance Rating)
PERF_SCORE_MAP = {
    "Consistently Exceeds": 2,
    "Exceeds": 1,
    "Meets": 0,
    "Inconsistently Meets": -1,
    "Needs Improvement": -2,
}


# -----------------------------
# COLUMN STANDARDIZATION
# -----------------------------
def _norm_col(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("/", " ").replace("-", " ")
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


EMP_COL_ALIASES = {
    "Employee ID": ["employee id", "employeeid", "emp id", "empid", "employee number", "employee#", "eid", "id"],
    "Job Family": ["job family", "jobfamily", "family", "job function", "function"],
    "Job Title": ["job title", "jobtitle", "title", "position", "position title"],
    "Job Level": ["job level", "joblevel", "level", "career level", "grade", "job grade", "pay grade", "band", "job band"],
    "Area Differential": ["area differential", "area", "geo", "geography", "geo differential", "location tier", "area diff"],
    "Base Pay": ["base pay", "base", "annual salary", "salary", "base salary"],
    "Bonus Paid": ["bonus paid", "bonus", "incentive", "incentive paid", "annual bonus", "bonus amount"],
    "Total Compensation": ["total compensation", "total comp", "total cash", "total rewards", "total pay", "total"],
    "Performance Rating": [
        "performance rating",
        "year end performance rating",
        "year-end performance rating",
        "year end rating",
        "year-end rating",
        "rating",
    ],
}


def standardize_employee_columns(employees: pd.DataFrame) -> pd.DataFrame:
    employees = employees.copy()
    employees.columns = [str(c).strip() for c in employees.columns]

    norm_to_actual = {_norm_col(c): c for c in employees.columns}
    rename_map = {}

    for required, aliases in EMP_COL_ALIASES.items():
        found_actual = None
        for a in [required] + aliases:
            key = _norm_col(a)
            if key in norm_to_actual:
                found_actual = norm_to_actual[key]
                break
        if found_actual:
            rename_map[found_actual] = required

    return employees.rename(columns=rename_map)


# -----------------------------
# HELPERS
# -----------------------------
def get_market_col(metric: str, pct: str) -> str:
    return f"{metric} {pct}"


def validate_employee_df(employees: pd.DataFrame):
    missing = [c for c in REQ_EMP_COLS if c not in employees.columns]
    if missing:
        raise ValueError(f"Employee sheet missing required columns: {missing}")


def validate_market_df(market: pd.DataFrame):
    if "Job Title" not in market.columns:
        raise ValueError("Market sheet must include 'Job Title'.")

    needed = []
    for pct in PCTS:
        needed += [
            get_market_col("Base Pay", pct),
            get_market_col("Bonus Paid", pct),
            get_market_col("Total Compensation", pct),
        ]
    missing = [c for c in needed if c not in market.columns]
    if missing:
        raise ValueError(
            "Market sheet is missing expected percentile columns.\n"
            f"Missing (sample): {missing[:10]}\n"
            "Expected columns like: 'Total Compensation P50', 'Bonus Paid P10', etc."
        )


def money(x):
    try:
        return "${:,.0f}".format(float(x))
    except Exception:
        return "—"


def uniq_sorted(s):
    ser = pd.Series(s).dropna().astype(str).map(lambda x: x.strip()).replace("", np.nan).dropna()
    return sorted(ser.unique().tolist())


def card(children, title=None):
    return html.Div(
        style={
            "background": CARD_BG,
            "border": f"1px solid {BORDER}",
            "borderRadius": "14px",
            "padding": "14px 14px",
            "boxShadow": "0 1px 0 rgba(0,0,0,0.03)",
        },
        children=[
            html.Div(title, style={"fontWeight": "800", "color": BRAND_DARK, "marginBottom": "8px"}) if title else None,
            children,
        ],
    )


def banner(text):
    return html.Div(
        text,
        style={
            "background": "#FFF1F2",
            "border": f"1px solid {BORDER}",
            "padding": "10px 12px",
            "borderRadius": "12px",
            "marginBottom": "12px",
            "fontSize": "13px",
            "color": BRAND_DARK,
        },
    )


def kpi_pill(label, value):
    return html.Div(
        style={
            "background": "#FFFFFF",
            "border": f"1px solid {BORDER}",
            "borderRadius": "999px",
            "padding": "8px 12px",
            "display": "flex",
            "gap": "8px",
            "alignItems": "baseline",
            "boxShadow": "0 1px 0 rgba(0,0,0,0.02)",
        },
        children=[
            html.Span(label, style={"color": BRAND_GRAY, "fontSize": "12px", "fontWeight": "700"}),
            html.Span(value, style={"color": BRAND_DARK, "fontSize": "14px", "fontWeight": "900"}),
        ],
    )


def normalize_area_diff(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if not s:
        return None
    sl = s.lower()
    if sl in ["high", "h"]:
        return "High"
    if sl in ["mid", "medium", "m"]:
        return "Mid"
    if sl in ["low", "l"]:
        return "Low"
    return s


def area_factor(area_value: str) -> float:
    a = normalize_area_diff(area_value)
    return AREA_FACTORS.get(a, 1.00)


# -----------------------------
# VARIABILITY INJECTION (salary-only)
# -----------------------------
def inject_employee_variability_from_market_with_area_and_perf(
    employees: pd.DataFrame,
    market: pd.DataFrame,
    *,
    join_key="Job Title",
    base_col="Base Pay",
    bonus_col="Bonus Paid",
    total_col="Total Compensation",
    area_col="Area Differential",
    perf_col="Performance Rating",
    perf_score_map=None,
    seed=13,
    beta_a=2.2,
    beta_b=2.2,
    area_effect_strength=0.25,
    perf_effect_strength=0.55,
    person_factor_sd=0.16,
    base_noise_sd=0.12,
    bonus_noise_sd=0.20,
    total_noise_sd=0.10,
    bonus_zero_prob=0.12,
):
    rng = np.random.default_rng(seed)
    emp = employees.copy()
    mkt = market.copy()

    if perf_score_map is None:
        perf_score_map = PERF_SCORE_MAP

    emp.columns = [str(c).strip() for c in emp.columns]
    mkt.columns = [str(c).strip() for c in mkt.columns]

    emp[join_key] = emp[join_key].astype(str).str.strip()
    mkt[join_key] = mkt[join_key].astype(str).str.strip()

    if area_col in emp.columns:
        emp[area_col] = emp[area_col].apply(normalize_area_diff)

    # numeric market percentiles
    for metric in ["Base Pay", "Bonus Paid", "Total Compensation"]:
        for pct in PCTS:
            col = f"{metric} {pct}"
            if col in mkt.columns:
                mkt[col] = pd.to_numeric(mkt[col], errors="coerce")

    mkt_cols = [join_key] + [
        f"{metric} {pct}"
        for metric in ["Base Pay", "Bonus Paid", "Total Compensation"]
        for pct in PCTS
    ]
    mkt_small = mkt[mkt_cols].drop_duplicates(subset=[join_key]).copy()
    emp2 = emp.merge(mkt_small, on=join_key, how="left")

    # apply area factor to market percentiles
    mf = emp2[area_col].apply(area_factor).astype(float) if area_col in emp2.columns else 1.0
    for metric in ["Base Pay", "Bonus Paid", "Total Compensation"]:
        for pct in PCTS:
            col = f"{metric} {pct}"
            if col in emp2.columns:
                emp2[col] = pd.to_numeric(emp2[col], errors="coerce") * mf

    # area score / perf score
    if area_col in emp2.columns:
        area_score = (
            emp2[area_col]
            .map({"High": 1.0, "Mid": 0.0, "Low": -1.0})
            .astype(float)
            .fillna(0.0)
            .to_numpy()
        )
    else:
        area_score = np.zeros(len(emp2), dtype=float)

    if perf_col in emp2.columns:
        perf_series = emp2[perf_col].astype(str).str.strip()
        perf_raw = perf_series.map(perf_score_map).astype(float).fillna(0.0).to_numpy()
    else:
        perf_raw = np.zeros(len(emp2), dtype=float)

    if np.std(perf_raw) > 1e-9:
        perf_score = (perf_raw - np.mean(perf_raw)) / (np.std(perf_raw) + 1e-9)
        perf_score = np.clip(perf_score, -2.0, 2.0)
    else:
        perf_score = np.zeros(len(emp2), dtype=float)

    u0 = rng.beta(beta_a, beta_b, size=len(emp2))
    u0 = np.clip(u0, 0.02, 0.98)

    def logit(p):
        p = np.clip(p, 1e-6, 1 - 1e-6)
        return np.log(p / (1 - p))

    def sigmoid(z):
        return 1 / (1 + np.exp(-z))

    z = logit(u0) + area_effect_strength * area_score + perf_effect_strength * perf_score + rng.normal(
        0, 0.12, size=len(emp2)
    )
    u = np.clip(sigmoid(z), 0.02, 0.98)

    pf = rng.normal(0, person_factor_sd, size=len(emp2))

    def piecewise_quantile(metric_name, i, uu):
        pcts_dict = {
            0.10: emp2.loc[i, f"{metric_name} P10"],
            0.25: emp2.loc[i, f"{metric_name} P25"],
            0.50: emp2.loc[i, f"{metric_name} P50"],
            0.75: emp2.loc[i, f"{metric_name} P75"],
            0.90: emp2.loc[i, f"{metric_name} P90"],
        }
        xs = np.array(sorted(pcts_dict.keys()), dtype=float)
        ys = np.array([pcts_dict[x] for x in xs], dtype=float)
        if np.any(~np.isfinite(ys)) or len(xs) < 2:
            return np.nan
        return float(np.interp(float(uu), xs, ys))

    def gen_component(metric_name, noise_sd, floor_zero=False, zero_prob=0.0):
        vals = []
        for i in range(len(emp2)):
            base = piecewise_quantile(metric_name, i, u[i])
            vals.append(base)
        vals = np.array(vals, dtype=float)
        eps = rng.normal(0, noise_sd, size=len(emp2))
        out = vals * np.exp(pf + eps)
        if floor_zero:
            zeros = rng.random(len(emp2)) < zero_prob
            out[zeros] = 0.0
            out = np.maximum(out, 0.0)
        return out

    base_like = gen_component("Base Pay", base_noise_sd, floor_zero=False)
    bonus = gen_component("Bonus Paid", bonus_noise_sd, floor_zero=True, zero_prob=bonus_zero_prob)
    total = gen_component("Total Compensation", total_noise_sd, floor_zero=False)

    emp2[base_col] = base_like
    emp2[bonus_col] = bonus
    emp2[total_col] = total

    # Drop temp market cols so later merges don't suffix/break
    drop_cols = [f"{metric} {pct}" for metric in ["Base Pay", "Bonus Paid", "Total Compensation"] for pct in PCTS]
    emp2 = emp2.drop(columns=[c for c in drop_cols if c in emp2.columns], errors="ignore")

    return emp2


# -----------------------------
# MODEL BUILD
# -----------------------------
def build_model_df(employees: pd.DataFrame, market: pd.DataFrame) -> pd.DataFrame:
    employees = employees.copy()
    market = market.copy()

    employees.columns = [str(c).strip() for c in employees.columns]
    market.columns = [str(c).strip() for c in market.columns]

    validate_employee_df(employees)
    validate_market_df(market)

    for col in ["Job Title", "Job Family", "Job Level", "Area Differential", "Performance Rating"]:
        employees[col] = employees[col].astype(str).str.strip()

    employees["Area Differential"] = employees["Area Differential"].apply(normalize_area_diff)

    for c in ["Base Pay", "Bonus Paid", "Total Compensation"]:
        employees[c] = pd.to_numeric(employees[c], errors="coerce")

    employees = employees.dropna(subset=["Job Title"]).copy()

    for pct in PCTS:
        for metric in ["Base Pay", "Bonus Paid", "Total Compensation"]:
            market[get_market_col(metric, pct)] = pd.to_numeric(market[get_market_col(metric, pct)], errors="coerce")

    df0 = employees.merge(market, on="Job Title", how="left")

    df0["Market Factor"] = df0["Area Differential"].apply(area_factor)

    for pct in PCTS:
        df0[f"Market Base Pay {pct}"] = df0[get_market_col("Base Pay", pct)] * df0["Market Factor"]
        df0[f"Market Bonus Paid {pct}"] = df0[get_market_col("Bonus Paid", pct)] * df0["Market Factor"]
        df0[f"Market Total Compensation {pct}"] = df0[get_market_col("Total Compensation", pct)] * df0["Market Factor"]

    return df0


def apply_filters(d, job_title, area):
    out = d.copy()
    if job_title != "All":
        out = out[out["Job Title"].astype(str) == str(job_title)]
    if area != "All":
        out = out[out["Area Differential"].astype(str) == str(area)]
    return out


def market_cols_for_metric(metric_value):
    if metric_value == "Base Pay":
        return [(pct, f"Market Base Pay {pct}") for pct in PCTS]
    if metric_value == "Bonus Paid":
        return [(pct, f"Market Bonus Paid {pct}") for pct in PCTS]
    if metric_value == "Total Compensation":
        return [(pct, f"Market Total Compensation {pct}") for pct in PCTS]
    raise ValueError("Unknown metric.")


# -----------------------------
# Breakout table (ALL Job Titles within a selected Job Family)
# -----------------------------
def build_job_family_breakout_table(d_all: pd.DataFrame, metric: str, sort_pct: str, sort_dir: str) -> go.Figure:
    if metric not in ["Base Pay", "Bonus Paid", "Total Compensation"]:
        return go.Figure().update_layout(template=PLOT_TEMPLATE, title="Unknown metric for breakout")

    if sort_pct not in PCTS:
        sort_pct = "P50"
    if sort_dir not in ["asc", "desc"]:
        sort_dir = "asc"

    if "Job Title" not in d_all.columns:
        return go.Figure().update_layout(template=PLOT_TEMPLATE, title="Missing column: Job Title")

    rows = []
    mcols = market_cols_for_metric(metric)

    # ✅ ALL jobs (no .head() cap)
    jobs = (
        d_all.groupby("Job Title")["Employee ID"]
        .count()
        .sort_values(ascending=False)
        .index
        .tolist()
    )

    for jt in jobs:
        dj = d_all[d_all["Job Title"].astype(str) == str(jt)]
        hc = int(dj["Employee ID"].count())
        if hc == 0:
            continue

        emp_p = {pct: float(dj[metric].quantile(PCT_Q[pct])) for pct in PCTS}
        mkt_p = {pct: float(dj[col].mean()) for pct, col in mcols}
        delta = {pct: emp_p[pct] - mkt_p[pct] for pct in PCTS}

        row = {"Job Title": str(jt), "Headcount": hc}
        for pct in PCTS:
            row[f"Δ{pct}"] = delta[pct]
        rows.append(row)

    out = pd.DataFrame(rows)
    if out.empty:
        return go.Figure().update_layout(
            template=PLOT_TEMPLATE,
            height=560,
            margin=dict(l=10, r=10, t=10, b=10),
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color=BRAND_DARK),
        )

    sort_col = f"Δ{sort_pct}"
    out = out.sort_values(sort_col, ascending=(sort_dir == "asc")).reset_index(drop=True)

    def fmt(x):
        try:
            return "${:,.0f}".format(float(x))
        except Exception:
            return "—"

    fig = go.Figure(
        data=[go.Table(
            header=dict(
                values=["Job Title", "Headcount"] + [f"Δ{p}" for p in PCTS],
                fill_color="#F9FAFB",
                align="left",
                font=dict(color=BRAND_DARK, size=12),
                line_color=BORDER
            ),
            cells=dict(
                values=[
                    out["Job Title"],
                    out["Headcount"],
                    out["ΔP10"].map(fmt),
                    out["ΔP25"].map(fmt),
                    out["ΔP50"].map(fmt),
                    out["ΔP75"].map(fmt),
                    out["ΔP90"].map(fmt),
                ],
                align="left",
                font=dict(color=BRAND_DARK, size=12),
                height=26,
                line_color=BORDER
            )
        )]
    )
    fig.update_layout(
        template=PLOT_TEMPLATE,
        height=560,
        margin=dict(l=10, r=10, t=40, b=10),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color=BRAND_DARK),
    )
    return fig


# -----------------------------
# LOAD DATA
# -----------------------------
def load_data_local_single_file():
    if not os.path.exists(EMPLOYEE_FILE):
        raise FileNotFoundError(f"Workbook not found: {EMPLOYEE_FILE}")

    xl = pd.ExcelFile(EMPLOYEE_FILE)
    sheets = xl.sheet_names

    if EMPLOYEE_SHEET not in sheets:
        raise ValueError(f"Employee sheet '{EMPLOYEE_SHEET}' not found. Available sheets: {sheets}")
    if MARKET_SHEET not in sheets:
        raise ValueError(f"Market sheet '{MARKET_SHEET}' not found. Available sheets: {sheets}")

    employees = pd.read_excel(EMPLOYEE_FILE, sheet_name=EMPLOYEE_SHEET)
    employees = standardize_employee_columns(employees)

    market = pd.read_excel(EMPLOYEE_FILE, sheet_name=MARKET_SHEET)
    market.columns = [str(c).strip() for c in market.columns]

    employees = inject_employee_variability_from_market_with_area_and_perf(
        employees,
        market,
        seed=13,
        perf_col="Performance Rating",
        perf_score_map=PERF_SCORE_MAP,
    )

    df_local = build_model_df(employees, market)

    banner_text = (
        f"LOCAL DATA: {os.path.basename(EMPLOYEE_FILE)} [{EMPLOYEE_SHEET}] + [{MARKET_SHEET}] | "
        "Market = National × Area Diff (High +12%, Mid 0%, Low -5%) | "
        "All roles treated as SALARY"
    )
    return df_local, banner_text


df, data_banner = load_data_local_single_file()


# -----------------------------
# DASH APP
# -----------------------------
app = Dash(__name__)
server = app.server
app.title = "Market Benchmarking Dashboard (Local)"


app.layout = html.Div(
    style={"fontFamily": "Arial", "background": BRAND_BG, "padding": "18px"},
    children=[
        html.Div(
            style={"display": "flex", "justifyContent": "space-between", "alignItems": "flex-end", "marginBottom": "10px"},
            children=[
                html.Div([
                    html.Div("Market Benchmarking", style={"fontSize": "22px", "fontWeight": "900", "color": BRAND_DARK}),
                    html.Div("Employee vs Market — National Market + Area Diff adjustment",
                             style={"fontSize": "13px", "color": BRAND_GRAY}),
                ]),
                html.Div("Benchmarking demo", style={"fontWeight": "900", "color": BRAND_RED}),
            ],
        ),
        banner(data_banner),

        # -----------------------------
        # TOP FILTERS (Metric, Job Title, Area Diff at far right)
        # -----------------------------
        card(
            title="Filters",
            children=html.Div(
                style={"display": "flex", "gap": "14px", "flexWrap": "wrap"},
                children=[
                    html.Div([
                        html.Label("Metric", style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"}),
                        dcc.Dropdown(
                            id="metric",
                            value="Total Compensation",
                            options=[
                                {"label": "Base Pay", "value": "Base Pay"},
                                {"label": "Bonus Paid", "value": "Bonus Paid"},
                                {"label": "Total Compensation", "value": "Total Compensation"},
                            ],
                            clearable=False,
                            style={"width": "260px"},
                        ),
                    ]),
                    html.Div([
                        html.Label("Job Title", style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"}),
                        dcc.Dropdown(
                            id="job_title",
                            value="All",
                            options=[],
                            clearable=False,
                            style={"width": "420px"},
                        ),
                    ]),
                    html.Div([
                        html.Label("Area Differential", style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"}),
                        dcc.Dropdown(
                            id="area",
                            value="All",
                            options=[{"label": "All", "value": "All"}] + [{"label": x, "value": x} for x in uniq_sorted(df["Area Differential"])],
                            clearable=False,
                            style={"width": "180px"},
                        ),
                    ]),
                ],
            ),
        ),

        html.Div(id="kpi_block", style={"marginTop": "12px"}),

        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1.25fr 0.75fr", "gap": "16px", "marginTop": "12px"},
            children=[
                card(dcc.Graph(id="pct_line", config={"displayModeBar": False}), title="Percentile Comparison"),
                card(
                    children=html.Div(
                        children=[
                            dcc.Graph(id="delta_table", config={"displayModeBar": False}),
                            html.Div(
                                id="delta_headcount_text",
                                style={"marginTop": "8px", "color": BRAND_DARK, "fontWeight": "900"},
                            ),
                        ]
                    ),
                    title="Employee − Market Delta (Table)",
                ),
            ],
        ),

        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px", "marginTop": "16px"},
            children=[
                card(dcc.Graph(id="dist_hist", config={"displayModeBar": False}), title="Employee Distribution vs Market Reference"),
                card(
                    title="Costing Scenarios",
                    children=html.Div(
                        children=[
                            html.Div(
                                style={"display": "flex", "gap": "12px", "flexWrap": "wrap", "alignItems": "end", "marginBottom": "8px"},
                                children=[
                                    html.Div(
                                        style={"minWidth": "360px"},
                                        children=[
                                            html.Label(
                                                "Year-End Performance Rating (Costing — Base Pay)",
                                                style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"},
                                            ),
                                            dcc.Dropdown(
                                                id="perf_filter",
                                                value=["All"],
                                                options=[{"label": "All", "value": "All"}]
                                                + [{"label": x, "value": x} for x in uniq_sorted(df["Performance Rating"])],
                                                multi=True,
                                                clearable=False,
                                            ),
                                            html.Div(
                                                "Select one or more ratings. If 'All' is selected, it overrides other selections.",
                                                style={"fontSize": "11px", "color": BRAND_GRAY, "marginTop": "6px", "fontWeight": "700"},
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                            dcc.Graph(id="cost_plot", config={"displayModeBar": False}),
                        ]
                    ),
                ),
            ],
        ),

        html.Div(style={"marginTop": "16px"}),
        card(
            title="Market Competitiveness Breakouts (Job Titles within a Job Family)",
            children=html.Div(
                children=[
                    html.Div(
                        style={"color": BRAND_GRAY, "fontSize": "12px", "marginBottom": "10px", "fontWeight": "700"},
                        children="These breakouts are independent of the top filters (Metric / Job Title / Area Differential).",
                    ),
                    html.Div(
                        style={"display": "flex", "gap": "12px", "flexWrap": "wrap", "marginBottom": "10px", "alignItems": "end"},
                        children=[
                            html.Div([
                                html.Label("Breakout Metric", style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"}),
                                dcc.Dropdown(
                                    id="breakout_metric",
                                    value="Total Compensation",
                                    options=[
                                        {"label": "Base Pay", "value": "Base Pay"},
                                        {"label": "Bonus Paid", "value": "Bonus Paid"},
                                        {"label": "Total Compensation", "value": "Total Compensation"},
                                    ],
                                    clearable=False,
                                    style={"width": "260px"},
                                ),
                            ]),
                            html.Div([
                                html.Label("Sort by Δ Percentile", style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"}),
                                dcc.Dropdown(
                                    id="breakout_sort_pct",
                                    value="P50",
                                    options=[{"label": p, "value": p} for p in PCTS],
                                    clearable=False,
                                    style={"width": "180px"},
                                ),
                            ]),
                            html.Div([
                                html.Label("Sort Order", style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"}),
                                dcc.Dropdown(
                                    id="breakout_sort_dir",
                                    value="asc",
                                    options=[
                                        {"label": "Smallest → Largest", "value": "asc"},
                                        {"label": "Largest → Smallest", "value": "desc"},
                                    ],
                                    clearable=False,
                                    style={"width": "220px"},
                                ),
                            ]),
                            html.Div([
                                html.Label("Job Family", style={"fontWeight": "800", "color": BRAND_DARK, "fontSize": "12px"}),
                                dcc.Dropdown(
                                    id="breakout_family",
                                    value="All",
                                    options=[{"label": "All", "value": "All"}] + [{"label": x, "value": x} for x in uniq_sorted(df["Job Family"])],
                                    clearable=False,
                                    style={"width": "320px"},
                                ),
                            ]),
                        ],
                    ),
                    dcc.Graph(id="breakout_table", config={"displayModeBar": False}),
                ]
            ),
        ),
    ],
)


# -----------------------------
# Callbacks
# -----------------------------
@app.callback(
    Output("job_title", "options"),
    Output("job_title", "value"),
    Input("area", "value"),
    Input("job_title", "value"),
)
def update_job_title_options(area, current_value):
    d = df.copy()
    if area != "All":
        d = d[d["Area Differential"].astype(str) == str(area)]

    titles = uniq_sorted(d["Job Title"])
    opts = [{"label": "All", "value": "All"}] + [{"label": t, "value": t} for t in titles]
    valid = set(["All"] + titles)
    return opts, (current_value if current_value in valid else "All")


@app.callback(
    Output("kpi_block", "children"),
    Output("pct_line", "figure"),
    Output("delta_table", "figure"),
    Output("delta_headcount_text", "children"),
    Output("dist_hist", "figure"),
    Output("cost_plot", "figure"),
    Input("metric", "value"),
    Input("job_title", "value"),
    Input("area", "value"),
    Input("perf_filter", "value"),
)
def update_main_dashboard(metric, job_title, area, perf_filter):
    d_all = apply_filters(df, job_title=job_title, area=area)

    empty = go.Figure().update_layout(template=PLOT_TEMPLATE)
    if d_all.empty:
        kpis = html.Div("No data under current filters.", style={"color": BRAND_DARK, "fontWeight": "800"})
        return kpis, empty, empty, "", empty, empty

    title_label = "All Job Titles" if (job_title is None or job_title == "All") else str(job_title)
    area_label = "All Areas" if (area is None or area == "All") else str(area)

    d = d_all.copy()

    mcols = market_cols_for_metric(metric)

    emp_p = {pct: float(d[metric].quantile(PCT_Q[pct])) for pct in PCTS}
    mkt_p = {pct: float(d[col].mean()) for pct, col in mcols}
    delta = {pct: emp_p[pct] - mkt_p[pct] for pct in PCTS}

    floor_p10 = d[[col for pct, col in mcols if pct == "P10"][0]]
    floor_p25 = d[[col for pct, col in mcols if pct == "P25"][0]]
    floor_p50 = d[[col for pct, col in mcols if pct == "P50"][0]]

    cost_to_p10 = float(np.maximum(0, floor_p10 - d[metric]).sum())
    cost_to_p25 = float(np.maximum(0, floor_p25 - d[metric]).sum())
    cost_to_p50 = float(np.maximum(0, floor_p50 - d[metric]).sum())

    headcount = int(len(d))
    pct_below_p10 = float((d[metric] < floor_p10).mean())
    pct_below_p25 = float((d[metric] < floor_p25).mean())
    pct_below_p50 = float((d[metric] < floor_p50).mean())

    kpis = html.Div(
        style={"display": "flex", "gap": "10px", "flexWrap": "wrap"},
        children=[
            kpi_pill("Job Title", title_label),
            kpi_pill("Area Differential", area_label),
            kpi_pill("Headcount", f"{headcount:,}"),
            kpi_pill(f"% < Mkt P10 ({metric})", f"{pct_below_p10*100:.1f}%"),
            kpi_pill(f"% < Mkt P25 ({metric})", f"{pct_below_p25*100:.1f}%"),
            kpi_pill(f"% < Mkt P50 ({metric})", f"{pct_below_p50*100:.1f}%"),
            kpi_pill("Cost to P10", money(cost_to_p10)),
            kpi_pill("Cost to P25", money(cost_to_p25)),
            kpi_pill("Cost to P50", money(cost_to_p50)),
        ]
    )

    x = PCTS
    fig_line = go.Figure()
    fig_line.add_trace(go.Scatter(
        x=x, y=[emp_p[p] for p in x], mode="lines+markers", name="Employee",
        line=dict(width=3), marker=dict(size=8)
    ))
    fig_line.add_trace(go.Scatter(
        x=x, y=[mkt_p[p] for p in x], mode="lines+markers", name="Market (Nat × Area)",
        line=dict(width=3, color=BRAND_RED), marker=dict(size=8, color=BRAND_RED)
    ))
    fig_line.update_layout(
        template=PLOT_TEMPLATE,
        title=f"Percentile Comparison — {metric} | {title_label} | {area_label}",
        yaxis_title=metric,
        height=420,
        margin=dict(l=50, r=20, t=60, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color=BRAND_DARK),
    )

    def fmt_money(v):
        try:
            return "${:,.0f}".format(float(v))
        except Exception:
            return "—"

    delta_df = pd.DataFrame({
        "Percentile": PCTS,
        "Employee": [emp_p[p] for p in PCTS],
        "Market": [mkt_p[p] for p in PCTS],
        "Delta (Emp - Mkt)": [delta[p] for p in PCTS],
    })

    fig_delta_table = go.Figure(
        data=[go.Table(
            header=dict(
                values=["Percentile", "Employee", "Market", "Delta (Emp − Mkt)"],
                fill_color="#F9FAFB",
                align="left",
                font=dict(color=BRAND_DARK, size=12),
                line_color=BORDER
            ),
            cells=dict(
                values=[
                    delta_df["Percentile"],
                    delta_df["Employee"].map(fmt_money),
                    delta_df["Market"].map(fmt_money),
                    delta_df["Delta (Emp - Mkt)"].map(fmt_money),
                ],
                align="left",
                height=28,
                font=dict(color=BRAND_DARK, size=12),
                line_color=BORDER
            )
        )]
    )
    fig_delta_table.update_layout(
        template=PLOT_TEMPLATE,
        height=420,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color=BRAND_DARK),
    )

    delta_headcount_text = f"Total Employee Headcount = {headcount:,}"

    fig_hist = px.histogram(
        d, x=metric, nbins=30,
        title=f"Employee Distribution — {metric} | {title_label} | {area_label}",
        template=PLOT_TEMPLATE
    )
    fig_hist.update_layout(
        height=420,
        margin=dict(l=50, r=20, t=60, b=50),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color=BRAND_DARK),
        xaxis_title=metric,
        yaxis_title="Employee Count",
        showlegend=False
    )
    for pct in PCTS:
        fig_hist.add_vline(
            x=mkt_p[pct],
            line_dash="dot",
            line_width=2,
            line_color=BRAND_RED,
            annotation_text=f"Mkt {pct}",
            annotation_position="top"
        )

    # Costing Scenarios
    d_cost = d.copy()
    selected_ratings = perf_filter if isinstance(perf_filter, list) else ["All"]
    if "All" in selected_ratings and len(selected_ratings) > 1:
        selected_ratings = ["All"]

    if metric == "Base Pay":
        if selected_ratings and "All" not in selected_ratings:
            d_cost = d_cost[d_cost["Performance Rating"].astype(str).isin([str(x) for x in selected_ratings])].copy()

        if d_cost.empty:
            fig_cost = go.Figure().update_layout(
                template=PLOT_TEMPLATE,
                title="Costing Scenarios — Base Pay | (No employees for selected performance rating filter)",
                height=420,
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color=BRAND_DARK),
            )
        else:
            mcols_cost = market_cols_for_metric("Base Pay")
            floor_p10_c = d_cost[[col for pct, col in mcols_cost if pct == "P10"][0]]
            floor_p25_c = d_cost[[col for pct, col in mcols_cost if pct == "P25"][0]]
            floor_p50_c = d_cost[[col for pct, col in mcols_cost if pct == "P50"][0]]

            cost_to_p10_c = float(np.maximum(0, floor_p10_c - d_cost["Base Pay"]).sum())
            cost_to_p25_c = float(np.maximum(0, floor_p25_c - d_cost["Base Pay"]).sum())
            cost_to_p50_c = float(np.maximum(0, floor_p50_c - d_cost["Base Pay"]).sum())

            fig_cost = px.bar(
                x=["Bring to P10", "Bring to P25", "Bring to P50"],
                y=[cost_to_p10_c, cost_to_p25_c, cost_to_p50_c],
                title=f"Costing Scenarios — Base Pay | {title_label} | {area_label}",
                labels={"x": "Scenario", "y": "Total Cost ($)"},
                template=PLOT_TEMPLATE
            )
            fig_cost.update_layout(
                height=420,
                margin=dict(l=50, r=20, t=60, b=50),
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color=BRAND_DARK),
            )
            fig_cost.update_traces(marker_color=BRAND_RED)
    else:
        fig_cost = px.bar(
            x=["Bring to P10", "Bring to P25", "Bring to P50"],
            y=[cost_to_p10, cost_to_p25, cost_to_p50],
            title=f"Costing Scenarios — {metric} | {title_label} | {area_label}",
            labels={"x": "Scenario", "y": "Total Cost ($)"},
            template=PLOT_TEMPLATE
        )
        fig_cost.update_layout(
            height=420,
            margin=dict(l=50, r=20, t=60, b=50),
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color=BRAND_DARK),
        )
        fig_cost.update_traces(marker_color=BRAND_RED)

    return kpis, fig_line, fig_delta_table, delta_headcount_text, fig_hist, fig_cost


@app.callback(
    Output("breakout_table", "figure"),
    Input("breakout_metric", "value"),
    Input("breakout_sort_pct", "value"),
    Input("breakout_sort_dir", "value"),
    Input("breakout_family", "value"),
)
def update_breakout_table(breakout_metric, breakout_sort_pct, breakout_sort_dir, breakout_family):
    d_all = df.copy()
    if breakout_family != "All":
        d_all = d_all[d_all["Job Family"].astype(str) == str(breakout_family)].copy()

    title_suffix = "All Job Families" if breakout_family == "All" else str(breakout_family)
    fig = build_job_family_breakout_table(d_all, breakout_metric, breakout_sort_pct, breakout_sort_dir)
    fig.update_layout(title=f"Job Titles Breakout — {breakout_metric} | {title_suffix}")
    return fig


# -----------------------------
# RUN
# -----------------------------
if __name__ == "__main__":
    print("Starting Dash server… open http://127.0.0.1:8050/ in your browser")
    app.run(debug=True)

       