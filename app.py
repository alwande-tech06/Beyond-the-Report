"""
APFA802 - Beyond the Report
ESG-adjusted financial risk dashboard for five JSE-listed mining companies, FY2021-FY2025.

Run:  streamlit run app.py
"""
from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy.cluster.hierarchy import fcluster, linkage
from scipy.stats import linregress

import esgrq as Q
import loader as L
import nlp as NLP

APP_DIR = Path(__file__).parent
ESG_FILE = APP_DIR / "data" / "esg_data.csv"
FEEDBACK_FILE = APP_DIR / "data" / "feedback.csv"

st.set_page_config(page_title="Beyond the Report - JSE mining", layout="wide")

COLOURS = {c: m["colour"] for c, m in L.COMPANIES.items()}
COMPANY_ORDER = list(L.COMPANIES)
FONT = "Poppins, Segoe UI, Arial, sans-serif"
NAVY = "#1F3B5A"
ORANGE = "#F39C12"

# Average USD/ZAR rates used only to convert Gold Fields' US$ amounts. Editable in the sidebar.
# Verify against the SARB / the companies' own reported average rates before submission.
FX_DEFAULT = {2021: 14.79, 2022: 16.36, 2023: 18.45, 2024: 18.33, 2025: 17.90}

st.markdown(
    f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap');
      html, body, [class*="css"], .stMarkdown, .stText {{ font-family: {FONT}; }}
      h1, h2, h3 {{ font-family: {FONT}; font-weight: 600; letter-spacing: -0.01em; color: {NAVY}; }}
      h1 {{ font-size: 1.9rem; margin-bottom: 0.1rem; }}
      h3 {{ font-size: 1.15rem; }}
      .lede {{ color: #5A6675; font-size: 1.0rem; max-width: 72ch; line-height: 1.55; }}

      /* Cards: metrics, charts, tables and bordered containers sit on white with a soft shadow */
      div[data-testid="stMetric"], div[data-testid="stPlotlyChart"], div[data-testid="stDataFrame"],
      div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: #FFFFFF; border-radius: 6px; box-shadow: 0 2px 10px rgba(31, 59, 90, 0.08);
      }}
      div[data-testid="stMetric"] {{ padding: 0.9rem 1rem; border-top: 3px solid {ORANGE}; }}
      div[data-testid="stPlotlyChart"] {{ padding: 0.6rem; }}
      div[data-testid="stMetricLabel"] p {{ color: #5A6675; font-weight: 500; }}
      div[data-testid="stMetricValue"] {{ font-size: 1.7rem; color: {NAVY}; }}
      /* First card in a row is the navy "hero" card */
      div[data-testid="stHorizontalBlock"] > div:first-child div[data-testid="stMetric"] {{
        background: {NAVY}; border-top-color: {NAVY};
      }}
      div[data-testid="stHorizontalBlock"] > div:first-child div[data-testid="stMetric"] * {{ color: #FFFFFF !important; }}

      /* Sidebar brand block */
      .brand {{ text-align: center; padding: 0.4rem 0 0.8rem; }}
      .brand .avatar {{ width: 76px; height: 76px; margin: 0 auto 0.7rem; border-radius: 50%; background: #FFFFFF;
                       display: flex; align-items: center; justify-content: center;
                       box-shadow: 0 0 0 6px rgba(255, 255, 255, 0.12); }}
      .brand .name {{ font-weight: 600; font-size: 1.15rem; letter-spacing: 0.06em; text-transform: uppercase; }}
      .brand .sub {{ font-size: 0.8rem; opacity: 0.75; }}
      section[data-testid="stSidebar"] div[role="radiogroup"] label {{ padding: 0.3rem 0.4rem; border-radius: 4px; }}
      section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{ background: rgba(255, 255, 255, 0.08); }}

      .insight {{ background: #FFFFFF; border-left: 4px solid {ORANGE}; border-radius: 4px; padding: 0.55rem 0.9rem;
                 margin: 0.45rem 0; max-width: 80ch; box-shadow: 0 1px 6px rgba(31, 59, 90, 0.06); }}
      .insight.warn {{ border-left-color: #C0392B; }}
      .tag {{ display:inline-block; padding: 0.1rem 0.6rem; border-radius: 10px; font-size: 0.8rem; font-weight: 500; }}
      .tag.low {{ background:#DDEBF7; color:{NAVY}; }}
      .tag.mod {{ background:#FDEBD0; color:#8A5A00; }}
      .tag.high {{ background:#F6D5D1; color:#8E2B20; }}
      .stButton button[kind="primary"], .stDownloadButton button {{ border-radius: 14px; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------------------
# Data
# --------------------------------------------------------------------------------------
@st.cache_data(show_spinner="Reading the company workbooks...")
def get_long() -> pd.DataFrame:
    return L.build_long()


@st.cache_data
def get_wide(fx_items: tuple | None) -> pd.DataFrame:
    return L.build_wide(get_long(), dict(fx_items) if fx_items else None)


def get_esg() -> pd.DataFrame:
    if "esg_df" in st.session_state:
        return st.session_state.esg_df
    if ESG_FILE.exists():
        try:
            df = L.load_esg(ESG_FILE)
            st.session_state.esg_df = df
            return df
        except Exception as e:  # noqa: BLE001
            st.sidebar.error(f"data/esg_data.csv could not be read: {e}")
    return pd.DataFrame(columns=L.ESG_COLUMNS)


def get_esgrq() -> dict | None:
    if "esgrq" in st.session_state:
        return st.session_state.esgrq
    if Q.WORKBOOK.exists():
        try:
            return Q.load(Q.WORKBOOK)
        except Exception as e:  # noqa: BLE001
            st.sidebar.error(f"data/esgrq_coding.xlsx could not be read: {e}")
    return None


def esgrq_template_bytes() -> bytes:
    buf = io.BytesIO()
    Q.write_template(buf, COMPANY_ORDER)
    return buf.getvalue()


def esg_template() -> pd.DataFrame:
    indicators = [
        ("E", "Scope 1 GHG emissions", "tCO2e", "lower", "Natural"),
        ("E", "Scope 2 GHG emissions", "tCO2e", "lower", "Natural"),
        ("E", "Energy consumption", "GJ", "lower", "Natural"),
        ("E", "Water withdrawal", "ML", "lower", "Natural"),
        ("E", "Significant environmental incidents (level 3+)", "count", "lower", "Natural"),
        ("S", "Fatalities", "count", "lower", "Human"),
        ("S", "Lost-time injury frequency rate (LTIFR)", "per million hours", "lower", "Human"),
        ("S", "Employees", "number", "n/a", "Human"),
        ("S", "Women in workforce", "%", "higher", "Human"),
        ("S", "Training spend", "R million", "higher", "Human"),
        ("S", "Socio-economic development spend", "R million", "higher", "Social and relationship"),
        ("G", "Independent directors on board", "%", "higher", "Social and relationship"),
        ("G", "Women on board", "%", "higher", "Social and relationship"),
        ("G", "ESG data externally assured (1 = yes, 0 = no)", "flag", "higher", "Social and relationship"),
    ]
    rows = [
        dict(company=c, fiscal_year=y, pillar=p, indicator=i, value="", unit=u, better_when=b, capital=cap,
             source_document="", report_page="")
        for c in COMPANY_ORDER for y in L.YEARS for p, i, u, b, cap in indicators
    ]
    return pd.DataFrame(rows, columns=L.ESG_COLUMNS)


def fmt_amount(v: float, currency: str) -> str:
    if pd.isna(v):
        return "–"
    sym = "US$" if currency == "USD" else "R"
    if abs(v) >= 1000:
        return f"{sym}{v / 1000:,.1f}bn"
    return f"{sym}{v:,.0f}m"


def fmt_metric(var: str, v: float) -> str:
    if pd.isna(v):
        return "–"
    unit = L.UNITS.get(var, "")
    if unit == "%":
        return f"{v:,.1f}%"
    if unit == "x":
        return f"{v:,.2f}x"
    return f"{v:,.0f}"


def style_fig(fig: go.Figure, height: int = 420, legend: bool = True) -> go.Figure:
    titled = bool(fig.layout.title.text)
    fig.update_layout(
        template="simple_white", height=height, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", font=dict(family=FONT, size=13, color="#1F2A37"),
        margin=dict(l=10, r=10, t=80 if titled and legend else 40, b=10), hoverlabel=dict(font_family=FONT),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title_text="") if legend else None,
        showlegend=legend,
    )
    if titled:
        fig.update_layout(title=dict(y=0.98, yanchor="top", font=dict(color=NAVY)))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E8ECF1", showgrid=True, zeroline=True, zerolinecolor="#B7C0CB", linecolor="#D9DEE5")
    fig.update_xaxes(linecolor="#D9DEE5")
    return fig


def band_tag(band) -> str:
    if pd.isna(band):
        return "–"
    cls = {"Low risk": "low", "Moderate risk": "mod", "High risk": "high"}[str(band)]
    return f'<span class="tag {cls}">{band}</span>'


# --------------------------------------------------------------------------------------
# Sidebar
# --------------------------------------------------------------------------------------
st.sidebar.markdown(
    f"""<div class="brand">
      <div class="avatar"><svg width="40" height="40" viewBox="0 0 24 24" fill="{NAVY}">
        <path d="M3 20h18v2H3zM5 10h3v8H5zm5.5-4h3v12h-3zM16 13h3v5h-3z"/></svg></div>
      <div class="name">Beyond the Report</div>
      <div class="sub">APFA802 · JSE mining · FY2021–FY2025</div>
    </div>""",
    unsafe_allow_html=True,
)

PAGES = ["Overview", "Reporting quality", "Trends", "Peer benchmark", "Risk scorecard", "ESG indicators",
         "Relationships", "Advanced analytics", "Company profile", "Data and quality", "Test the dashboard"]
page = st.sidebar.radio("Go to", PAGES, label_visibility="collapsed")

st.sidebar.divider()
year = st.sidebar.select_slider("Financial year", options=L.YEARS, value=2025)
companies = st.sidebar.multiselect("Companies", COMPANY_ORDER, default=COMPANY_ORDER)
if not companies:
    st.sidebar.warning("Select at least one company.")
    companies = COMPANY_ORDER

convert = st.sidebar.toggle("Show Gold Fields amounts in rand", value=True,
                            help="Gold Fields reports in US dollars. Ratios are unaffected by this setting.")
fx = dict(FX_DEFAULT)
if convert:
    with st.sidebar.expander("Average USD/ZAR rates"):
        st.caption("Check these against SARB averages or Gold Fields' reported rates.")
        for y in L.YEARS:
            fx[y] = st.number_input(str(y), value=FX_DEFAULT[y], step=0.01, format="%.2f", key=f"fx{y}")

wide_all = get_wide(tuple(sorted(fx.items())) if convert else None)
long_df = get_long()
esg = get_esg()
esg_sc = L.esg_scores(esg) if not esg.empty else None

if "weights" not in st.session_state:
    st.session_state.weights = {"Profitability": 25, "Liquidity": 15, "Solvency": 25,
                                "Cash generation": 25, "Stability": 10, "ESG": 0}
score_all = L.scorecard(wide_all, esg_sc, st.session_state.weights)

wide = wide_all[wide_all.company.isin(companies)]
score = score_all[score_all.company.isin(companies)]
cur = wide[wide.fiscal_year == year].set_index("company").reindex([c for c in COMPANY_ORDER if c in companies])
cur_sc = score[score.fiscal_year == year].set_index("company").reindex(cur.index)

st.sidebar.divider()
st.sidebar.caption(
    "Impala's year ends 30 June; the others 31 December. "
    + ("ESG data loaded." if not esg.empty else "No ESG data loaded yet; see ESG indicators.")
)


# ======================================================================================
# Pages
# ======================================================================================
def page_overview():
    st.title("Beyond the Report")
    st.markdown(
        '<p class="lede">How exposed are South Africa\'s listed miners to financial risk, and does the information '
        "in their integrated reports make that risk visible? This dashboard turns five years of reported numbers for "
        "five JSE mining companies into comparable ratios, a risk score and data-quality checks.</p>",
        unsafe_allow_html=True,
    )

    cols = st.columns(len(cur.index))
    for col, c in zip(cols, cur.index):
        r, s = cur.loc[c], cur_sc.loc[c]
        prev = score_all[(score_all.company == c) & (score_all.fiscal_year == year - 1)]
        delta = None if prev.empty or pd.isna(s.Composite) else f"{s.Composite - prev.Composite.iloc[0]:+.1f} vs FY{year - 1}"
        with col:
            st.metric(c, f"{s.Composite:.0f} / 100" if pd.notna(s.Composite) else "–", delta)
            st.markdown(band_tag(s["Risk band"]), unsafe_allow_html=True)

    st.subheader(f"FY{year} at a glance")
    table = pd.DataFrame({
        "Revenue": [fmt_amount(cur.loc[c, "revenue"], cur.loc[c, "currency"]) for c in cur.index],
        "Net margin": [fmt_metric("net_margin", v) for v in cur.net_margin],
        "ROE": [fmt_metric("roe", v) for v in cur.roe],
        "Current ratio": [fmt_metric("current_ratio", v) for v in cur.current_ratio],
        "Debt / equity": [fmt_metric("debt_to_equity", v) for v in cur.debt_to_equity],
        "Interest cover": [fmt_metric("interest_cover", v) for v in cur.interest_cover],
        "FCF margin": [fmt_metric("fcf_margin", v) for v in cur.fcf_margin],
        "Risk score": [f"{v:.0f}" if pd.notna(v) else "–" for v in cur_sc.Composite],
    }, index=cur.index)
    table.index.name = "Company"
    st.dataframe(table, width="stretch")

    st.subheader("What the numbers say")
    for text, warn in insights():
        st.markdown(f'<div class="insight{" warn" if warn else ""}">{text}</div>', unsafe_allow_html=True)

    st.subheader("Risk score over time")
    fig = px.line(score, x="fiscal_year", y="Composite", color="company", markers=True,
                  color_discrete_map=COLOURS, category_orders={"company": COMPANY_ORDER},
                  labels={"fiscal_year": "Financial year", "Composite": "Risk score (100 = lowest risk)"})
    fig.add_hrect(y0=0, y1=40, fillcolor="#F6D5D1", opacity=0.35, line_width=0)
    fig.add_hrect(y0=65, y1=100, fillcolor="#DDEBF7", opacity=0.45, line_width=0)
    fig.update_xaxes(dtick=1)
    fig.update_yaxes(range=[0, 100])
    st.plotly_chart(style_fig(fig), width="stretch")
    st.caption("Shaded bands: below 40 = high risk, above 65 = low risk. Weights are set on the Risk scorecard page.")


def insights() -> list[tuple[str, bool]]:
    out = []
    c = cur.dropna(subset=["net_margin"])
    if not c.empty:
        best = c.net_margin.idxmax()
        out.append((f"<b>{best}</b> converted the most revenue into profit in FY{year}, "
                    f"with a net margin of {c.net_margin.max():.1f}%.", False))
    losers = cur[cur.net_profit < 0].index.tolist()
    if losers:
        out.append((f"<b>{', '.join(losers)}</b> reported a loss in FY{year}. "
                    "Check the impairment notes: large write-downs of mining assets usually drive this.", True))
    lev = cur.dropna(subset=["debt_to_equity"])
    if not lev.empty and lev.debt_to_equity.max() > 0.5:
        hi = lev.debt_to_equity.idxmax()
        out.append((f"<b>{hi}</b> carries the most debt relative to equity ({lev.debt_to_equity.max():.2f}x), "
                    "which leaves less room to absorb a commodity-price fall or a climate-related cost shock.", True))
    liq = cur.dropna(subset=["current_ratio"])
    if not liq.empty and liq.current_ratio.min() < 1.3:
        lo = liq.current_ratio.idxmin()
        out.append((f"<b>{lo}</b> has the thinnest short-term liquidity (current ratio {liq.current_ratio.min():.2f}x).", True))
    first = score_all[score_all.fiscal_year == L.YEARS[0]].set_index("company").Composite
    now = score_all[score_all.fiscal_year == year].set_index("company").Composite
    ch = (now - first).reindex(cur.index).dropna()
    if year > L.YEARS[0] and not ch.empty:
        w = ch.idxmin()
        out.append((f"Since FY{L.YEARS[0]}, <b>{w}</b>'s risk score has moved the most in the wrong direction "
                    f"({ch.min():+.0f} points)."
                    + (" The PGM price fall after 2021 shows up across the platinum producers." if "PGM" in L.COMPANIES[w]["commodity"] else ""),
                    ch.min() < -15))
    if esg.empty:
        out.append(("The score currently uses financial information only. Load ESG indicators to see whether "
                    "sustainability performance changes the ranking.", False))
    return out


BAND_CLS = {"Strong": "low", "Moderate": "mod", "Weak": "high"}
SEQ = [[0, "#FAD7A0"], [0.5, "#EEF1F5"], [1, "#6E9CC8"]]


def heat(z: pd.DataFrame, text: pd.DataFrame | None = None, zmax: float = 100, height: int = 300,
         scale: list | None = None) -> go.Figure:
    t = text if text is not None else z.round(0).map(lambda v: "" if pd.isna(v) else f"{v:.0f}")
    fig = go.Figure(go.Heatmap(z=z.values, x=list(z.columns), y=list(z.index), zmin=0, zmax=zmax, text=t.values,
                               texttemplate="%{text}", colorscale=scale or SEQ, showscale=False, xgap=2, ygap=2,
                               hoverongaps=False))
    fig.update_yaxes(autorange="reversed")
    fig.update_xaxes(side="top")
    return style_fig(fig, height, legend=False)


def page_esgrq():
    st.title("ESG reporting quality")
    st.markdown(
        '<p class="lede">Is the ESG information in these reports comparable, material, financially connected and '
        "credible? Each company's reporting suite is coded on 17 disclosure items (0–4), grouped into six dimensions. "
        "Every score links back to the report and page it came from.</p>", unsafe_allow_html=True)

    data = get_esgrq()
    with st.expander("Coding workbook", expanded=data is None):
        c1, c2 = st.columns(2)
        c1.markdown("The group codes the reports in one Excel workbook (sheets: Coding, Assurance, Comparability). "
                    "It is saved at <code>data/esgrq_coding.xlsx</code> and loads automatically; you can also "
                    "upload a copy here to preview it.", unsafe_allow_html=True)
        c1.download_button("Download blank coding workbook", esgrq_template_bytes(), "esgrq_coding.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        up = c2.file_uploader("Upload completed workbook", type="xlsx", key="esgrq_up")
        if up is not None:
            try:
                st.session_state.esgrq = Q.load(up)
                c2.success("Workbook loaded for this session.")
                data = st.session_state.esgrq
            except Exception as e:  # noqa: BLE001
                c2.error(f"Could not read the workbook: {e}")
        with st.container():
            st.caption("Scoring scale: " + " · ".join(f"{k} = {v}" for k, v in Q.SCALE.items()))

    if data is None:
        st.info("No coding workbook found. Download the blank workbook above, fill it in and save it as "
                "data/esgrq_coding.xlsx.")
        return

    coding = data["Coding"][data["Coding"].company.isin(companies)]
    assurance = data["Assurance"][data["Assurance"].company.isin(companies)]
    comp_in = data["Comparability"]
    years = sorted(coding.report_year.dropna().unique().tolist()) or [2025]
    ry = years[-1] if len(years) == 1 else st.select_slider("Report year", years, value=years[-1])
    cy = coding[coding.report_year == ry]
    order = [c for c in COMPANY_ORDER if c in companies]

    v = Q.validation(cy)
    k = st.columns(4)
    k[0].metric("Items coded", f"{v['scored']} / {v['expected']}")
    k[1].metric("With source and page", f"{v['traceable']} / {v['scored']}" if v["scored"] else "–")
    comp_tbl, comp_score = Q.comparability(comp_in[comp_in.report_year == ry], order)
    k[2].metric("Comparability score", f"{comp_score:.0f}%" if comp_score is not None else "–",
                help="Share of common indicators meeting every comparability criterion (paper section 9.3).")
    k[3].metric("Second-coder agreement", f"{v['exact']:.0f}%" if v["rechecked"] else "–",
                help="Exact agreement between score and score_recheck.")

    tabs = st.tabs(["Scorecard", "Comparability", "Materiality & connectivity", "Assurance", "Evidence trail",
                    "Validation"])

    # ---- Scorecard -------------------------------------------------------------------
    with tabs[0]:
        sc = Q.composite(cy)
        if sc.empty:
            st.info("No items scored yet. Fill in the score column on the Coding sheet.")
        else:
            sc = sc.set_index("company").reindex(order)
            cols = st.columns(len(order))
            for col, c in zip(cols, order):
                r = sc.loc[c]
                with col:
                    st.metric(c, f"{r.ESGRQ:.0f} / 100" if pd.notna(r.ESGRQ) else "–")
                    if pd.notna(r.ESGRQ):
                        st.markdown(f'<span class="tag {BAND_CLS[r.Band]}">{r.Band}</span> '
                                    f'<small>{r.Coverage:.0f}% coded</small>', unsafe_allow_html=True)
            st.subheader("Dimension profile")
            st.caption("0–100 = mean item score ÷ 4 × 100. Blank = no item in that dimension coded yet.")
            st.plotly_chart(heat(sc[Q.DIMENSIONS].dropna(how="all")), width="stretch")
            if (sc.Coverage < 100).any():
                st.markdown('<div class="insight warn">Some companies are only partly coded. Scores cover coded items '
                            "only, so compare them once coverage reaches 100%.</div>", unsafe_allow_html=True)

            st.subheader("Item scores")
            items = cy.pivot_table(index="company", columns="code", values="score").reindex(order)
            items = items.reindex(columns=[c for c in Q.ITEM_TABLE.code if c in items.columns])
            st.plotly_chart(heat(items, zmax=4, height=280), width="stretch")
            st.caption("Hover shows the raw 0–4 score. Item names: " + ", ".join(
                f"{r.code} {r.indicator}" for r in Q.ITEM_TABLE.itertuples()))

            st.subheader("Six capitals coverage")
            cap = Q.capital_scores(cy).pivot_table(index="company", columns="capital", values="score").reindex(order)
            caps = ["Financial", "Manufactured", "Intellectual", "Human", "Social and relationship", "Natural"]
            st.plotly_chart(heat(cap.reindex(columns=caps)), width="stretch")
            st.caption("Manufactured and intellectual capital have no item in the Appendix A framework, so they stay "
                       "blank. That is a limitation of the coding framework, not of the companies.")

            st.subheader("Reporting quality against financial risk")
            fin = score_all[score_all.fiscal_year == ry][["company", "Composite"]]
            m = sc.reset_index()[["company", "ESGRQ"]].merge(fin, on="company").dropna()
            m["fiscal_year"] = ry
            if len(m) >= 3:
                scatter_with_fit(m, "ESGRQ", "Composite", "ESGRQ (0–100)", f"FY{ry} financial risk score (100 = lowest risk)")
            else:
                st.info("Needs ESGRQ scores for at least three companies.")

    # ---- Comparability ---------------------------------------------------------------
    with tabs[1]:
        st.caption("For each common indicator, the Comparability sheet records how each company reports it. An "
                   "indicator is comparable only if all companies report it with the same unit, boundary and period "
                   "(and rate basis, where relevant).")
        if comp_tbl.empty:
            st.info("Fill in the Comparability sheet (at least the 'reported' column) to see this analysis.")
        else:
            show = comp_tbl.copy()
            for c in Q.CRITERIA:
                show[c] = show[c].map({True: "✓", False: "✗", None: "n/a"})
            show["Reporting"] = show.companies_reporting.astype(str) + " / " + show["of"].astype(str)
            show = show.rename(columns=Q.CRITERIA | {"indicator": "Indicator", "status": "Status"})
            st.dataframe(show[["Indicator", "Reporting", *Q.CRITERIA.values(), "Status"]], hide_index=True,
                         width="stretch")
            ind = st.selectbox("How each company reports", comp_tbl.indicator)
            d = comp_in[(comp_in.indicator == ind) & (comp_in.report_year == ry) & comp_in.company.isin(order)]
            st.dataframe(d.drop(columns=["indicator", "report_year"]).assign(
                reported=d.reported.map({1.0: "Yes", 0.0: "No"})), hide_index=True, width="stretch")

    # ---- Materiality & connectivity --------------------------------------------------
    with tabs[2]:
        mat = cy.pivot_table(index="company", columns="code", values="material").reindex(order)
        if mat.empty or mat.isna().all().all():
            st.info("Fill in the 'material' column on the Coding sheet.")
        else:
            mat = mat.reindex(columns=[c for c in Q.ITEM_TABLE.code if c in mat.columns])
            txt = mat.map(lambda x: "" if pd.isna(x) else ("Yes" if x else "No"))
            st.subheader("Material-topic map")
            st.plotly_chart(heat(mat, text=txt, zmax=1, height=280), width="stretch")
            share = mat.mean()
            names = Q.ITEM_TABLE.set_index("code").indicator
            common = [names[c] for c in share.index if share[c] == 1]
            some = [names[c] for c in share.index if 0 < share[c] < 1]
            st.markdown(f"**Material for every company:** {', '.join(common) or 'none'}  \n"
                        f"**Company-specific:** {', '.join(some) or 'none'}")
        st.subheader("Financial connectivity")
        conn = Q.connectivity(cy).set_index("company").reindex(order)
        f1 = cy[cy.code == "F1"].set_index("company").score.reindex(order)
        t = pd.DataFrame({"F1 score (0–4)": f1, "E/S/G items with a stated financial effect (%)": pd.to_numeric(conn.linked_share).round(0)})
        st.dataframe(t, width="stretch")
        linked = cy[(cy.financial_link == 1)][["company", "code", "indicator", "evidence_summary", "source_document",
                                               "report_page"]]
        if not linked.empty:
            with st.expander(f"Evidence of financial links ({len(linked)})"):
                st.dataframe(linked, hide_index=True, width="stretch")

    # ---- Assurance -------------------------------------------------------------------
    with tabs[3]:
        a = assurance[assurance.report_year == ry].set_index("company").reindex(order)
        filled = a[["provider", "standard", "level", "scope"]].fillna("").ne("").any(axis=1)
        a1 = cy[cy.code == "A1"].set_index("company").score.reindex(order)
        cov = cy[cy.assured.notna()].groupby("company").assured.mean().mul(100).reindex(order)
        tbl = a[["provider", "standard", "level", "scope", "subject_matter", "conclusion", "source_document",
                 "report_page"]].copy()
        tbl.insert(0, "A1 score", a1)
        tbl.insert(1, "Coded items inside assurance scope (%)", cov.round(0))
        st.dataframe(tbl, width="stretch")
        if not filled.all():
            st.markdown('<div class="insight warn">Assurance details still missing for: '
                        + ", ".join(a.index[~filled.values]) + ". The notes column holds what the draft paper "
                        "claims; confirm each against the independent assurance report.</div>", unsafe_allow_html=True)
            with st.expander("Draft paper claims to verify"):
                st.dataframe(a[["notes"]], width="stretch")
        st.caption("An 'assured' label alone is not enough: compare what was assured, to what level, under which "
                   "standard and with what conclusion.")

    # ---- Evidence trail --------------------------------------------------------------
    with tabs[4]:
        c1, c2 = st.columns(2)
        co = c1.selectbox("Company", order, key="ev_co")
        code = c2.selectbox("Item", Q.ITEM_TABLE.code, key="ev_item",
                            format_func=lambda x: f"{x} {Q.ITEM_TABLE.set_index('code').indicator[x]}")
        row = cy[(cy.company == co) & (cy.code == code)]
        if row.empty or pd.isna(row.score.iloc[0]):
            st.info("Not coded yet.")
        else:
            r = row.iloc[0]
            st.metric(f"{r.indicator} score", f"{int(r.score)} / 4", Q.SCALE[int(r.score)], delta_color="off")
            src = f"{r.source_document}, p. {r.report_page}" if r.source_document and r.report_page else \
                "⚠ source or page missing"
            st.markdown(f"**Evidence:** {r.evidence_summary or '–'}  \n**Source:** {src}  \n"
                        f"**Coded by:** {r.coder or '–'}" + (f" · **Recheck:** {int(r.score_recheck)} by {r.recheck_by}"
                                                              if pd.notna(r.score_recheck) else ""))
        full = cy[["company", "code", "indicator", "score", "material", "financial_link", "assured",
                   "evidence_summary", "source_document", "report_page", "coder"]]
        with st.expander("All coded evidence"):
            st.dataframe(full, hide_index=True, width="stretch")
        st.download_button("Download evidence table (CSV)", full.to_csv(index=False).encode(),
                           f"esgrq_evidence_{ry}.csv", "text/csv")

    # ---- Validation ------------------------------------------------------------------
    with tabs[5]:
        st.caption("Validation measures from section 13.3 of the paper. Report these numbers only once coding and "
                   "rechecking are complete.")
        k = st.columns(4)
        k[0].metric("Rows rechecked", v["rechecked"])
        k[1].metric("Exact agreement", f"{v['exact']:.0f}%" if v["rechecked"] else "–")
        k[2].metric("Within ±1 point", f"{v['within1']:.0f}%" if v["rechecked"] else "–")
        k[3].metric("Weighted kappa", f"{v['kappa']:.2f}" if pd.notna(v["kappa"]) else "–",
                    help="Quadratic-weighted Cohen's kappa. Landis & Koch (1977): 0.61–0.80 substantial, "
                         "above 0.80 almost perfect.")
        checks = [
            ("Traceability: scored items without source or page", v["unsourced"]),
            ("Data integrity: scores outside 0–4", v["invalid"]),
            ("Data integrity: duplicate company/year/item rows", v["duplicates"]),
            ("Data integrity: unknown item codes", v["unknown"]),
            ("Scoring accuracy: coder disagreements", v["disagreements"]),
        ]
        for label, df in checks:
            if df.empty:
                st.markdown(f"✓ {label}: none")
            else:
                with st.expander(f"✗ {label}: {len(df)}"):
                    st.dataframe(df[["company", "code", "indicator", "score", "score_recheck", "source_document",
                                     "report_page"]], hide_index=True, width="stretch")
        st.markdown("**Calculation consistency.** Scores are recomputed from the workbook on every load with a fixed "
                    "formula (dimension = mean item score × 25; ESGRQ = mean of dimensions), so identical input "
                    "always gives identical output. To check by hand, recompute one company's scores in Excel and "
                    "compare them with the Scorecard tab.")


def page_trends():
    st.title("Trends")
    options = [v for v in L.RATIOS + ["revenue", "net_profit", "operating_cash_flow", "free_cash_flow", "capex",
                                      "total_debt", "net_debt", "dividends_paid", "total_assets"]]
    c1, c2 = st.columns([3, 2])
    var = c1.selectbox("Measure", options, format_func=lambda v: L.LABELS.get(v, v))
    is_amount = var not in L.RATIOS
    index = c2.toggle("Index to FY2021 = 100", value=False, disabled=not is_amount,
                      help="Puts companies of different size and currency on the same scale.")

    d = wide[["company", "fiscal_year", var, "currency"]].copy()
    ylab = L.LABELS.get(var, var)
    if is_amount and index:
        base = d[d.fiscal_year == L.YEARS[0]].set_index("company")[var]
        d[var] = d.apply(lambda r: r[var] / base.get(r.company) * 100 if base.get(r.company) else np.nan, axis=1)
        ylab += " (FY2021 = 100)"
    elif is_amount:
        ylab += " (R million)" if convert else " (million, reporting currency)"
    elif L.UNITS.get(var) == "%":
        ylab += " (%)"

    fig = px.line(d, x="fiscal_year", y=var, color="company", markers=True, color_discrete_map=COLOURS,
                  category_orders={"company": COMPANY_ORDER}, labels={"fiscal_year": "Financial year", var: ylab})
    fig.update_xaxes(dtick=1)
    st.plotly_chart(style_fig(fig, 460), width="stretch")
    row = L.DICT_DF.set_index("variable").loc[var]
    st.caption(f"**{row.label}**: {row.definition}")
    if is_amount and not convert and "Gold Fields" in companies and not index:
        st.info("Gold Fields is shown in US$ million. Turn on the rand toggle or the index view to compare size.")

    st.subheader("Change over the period")
    first, last = L.YEARS[0], year
    t = wide.pivot(index="company", columns="fiscal_year", values=var)
    if first in t and last in t:
        chg = pd.DataFrame({f"FY{first}": t[first].map(lambda v: fmt_metric(var, v) if not is_amount else f"{v:,.0f}"),
                            f"FY{last}": t[last].map(lambda v: fmt_metric(var, v) if not is_amount else f"{v:,.0f}")})
        if is_amount:
            chg["Change"] = ((t[last] / t[first] - 1) * 100).map(lambda v: f"{v:+.0f}%" if pd.notna(v) else "–")
        else:
            chg["Change"] = (t[last] - t[first]).map(lambda v: f"{v:+.1f}" if pd.notna(v) else "–")
        st.dataframe(chg, width="stretch")


def page_benchmark():
    st.title("Peer benchmark")
    st.markdown(f'<p class="lede">FY{year}. Colour shows each company\'s rank within the selected peers on every ratio, '
                "with direction taken into account (green = stronger). The number in each cell is the actual value.</p>",
                unsafe_allow_html=True)
    ratio_list = ["net_margin", "ebit_margin", "gross_margin", "roa", "roe", "current_ratio", "quick_ratio",
                  "debt_to_equity", "liabilities_to_assets", "interest_cover", "ocf_margin", "fcf_margin",
                  "revenue_growth"]
    better = L.DICT_DF.set_index("variable").better_when
    z, text = [], []
    for v in ratio_list:
        s = cur[v]
        rk = s.rank(pct=True)
        if better.get(v) == "lower":
            rk = 1 - rk + 1 / max(s.notna().sum(), 1)
        z.append(rk.values)
        text.append([fmt_metric(v, x) for x in s.values])
    fig = go.Figure(go.Heatmap(
        z=z, x=list(cur.index), y=[L.LABELS[v] for v in ratio_list], text=text, texttemplate="%{text}",
        colorscale=[[0, "#F5B041"], [0.5, "#EEF1F5"], [1, "#6E9CC8"]], showscale=False, xgap=2, ygap=2,
        hovertemplate="%{x}<br>%{y}: %{text}<extra></extra>",
    ))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    st.plotly_chart(style_fig(fig, 560, legend=False), width="stretch")

    st.subheader("Against the peer median")
    var = st.selectbox("Ratio", ratio_list, format_func=lambda v: L.LABELS[v])
    d = cur[[var]].reset_index()
    med = d[var].median()
    fig = px.bar(d, x="company", y=var, color="company", color_discrete_map=COLOURS, text=d[var].map(lambda x: fmt_metric(var, x)),
                 labels={"company": "", var: L.LABELS[var]})
    fig.add_hline(y=med, line_dash="dot", line_color="#1E2328", annotation_text=f"Peer median {fmt_metric(var, med)}",
                  annotation_position="top left")
    st.plotly_chart(style_fig(fig, 380, legend=False), width="stretch")


def page_scorecard():
    st.title("Risk scorecard")
    st.markdown('<p class="lede">Each pillar is scored 0–100 across all company-years in the sample, so a firm can be '
                "compared with its peers and with its own past. 100 means the strongest position in the sample. "
                "Change the weights to reflect a particular stakeholder's view.</p>", unsafe_allow_html=True)

    PRESETS = {
        "Balanced": {"Profitability": 25, "Liquidity": 15, "Solvency": 25, "Cash generation": 25, "Stability": 10, "ESG": 0},
        "Lender": {"Profitability": 10, "Liquidity": 30, "Solvency": 35, "Cash generation": 20, "Stability": 5, "ESG": 0},
        "Equity investor": {"Profitability": 35, "Liquidity": 10, "Solvency": 15, "Cash generation": 30, "Stability": 10, "ESG": 0},
        "Sustainability-focused investor": {"Profitability": 15, "Liquidity": 10, "Solvency": 15, "Cash generation": 15,
                                            "Stability": 5, "ESG": 40},
    }
    names = ["Profitability", "Liquidity", "Solvency", "Cash generation", "Stability", "ESG"]
    for p in names:
        st.session_state.setdefault(f"w_{p}", st.session_state.weights.get(p, 0))
    if esg.empty:
        st.session_state["w_ESG"] = 0

    def apply(preset):
        for p, v in PRESETS[preset].items():
            st.session_state[f"w_{p}"] = 0 if (p == "ESG" and esg.empty) else v

    st.caption("Weight presets")
    bcols = st.columns(len(PRESETS))
    for col, name in zip(bcols, PRESETS):
        col.button(name, on_click=apply, args=(name,), width="stretch",
                   disabled=(name == "Sustainability-focused investor" and esg.empty),
                   help="Needs ESG data" if name == "Sustainability-focused investor" and esg.empty else None)
    cols = st.columns(6)
    for col, p in zip(cols, names):
        col.slider(p, 0, 50, step=5, key=f"w_{p}", disabled=(p == "ESG" and esg.empty),
                   help="Load ESG data on the ESG indicators page to use this pillar." if p == "ESG" and esg.empty else None)
    new = {p: st.session_state[f"w_{p}"] for p in names}
    if new != st.session_state.weights:
        st.session_state.weights = new
        st.rerun()

    pillars = [p for p in ["Profitability", "Liquidity", "Solvency", "Cash generation", "Stability", "ESG"] if p in cur_sc]
    st.subheader(f"FY{year} ranking")
    t = cur_sc[pillars + ["Composite", "Risk band"]].sort_values("Composite", ascending=False)
    t.index.name = "Company"
    st.dataframe(t.round(0), width="stretch",
                 column_config={**{p: st.column_config.NumberColumn(p, format="%.0f") for p in pillars},
                                "Composite": st.column_config.ProgressColumn("Composite", min_value=0, max_value=100,
                                                                             format="%.0f", width="medium")})

    left, right = st.columns(2)
    with left:
        st.subheader("Pillar profile")
        fig = go.Figure()
        for c in cur_sc.index:
            vals = cur_sc.loc[c, pillars].tolist()
            fig.add_trace(go.Scatterpolar(r=vals + vals[:1], theta=pillars + pillars[:1], name=c,
                                          line=dict(color=COLOURS[c], width=2), fill="none"))
        fig.update_layout(polar=dict(radialaxis=dict(range=[0, 100], showticklabels=False, gridcolor="#E3E6E2"),
                                     domain=dict(x=[0.08, 0.92], y=[0.02, 0.85])))
        st.plotly_chart(style_fig(fig, 460), width="stretch")
    with right:
        st.subheader("How the score moves")
        focus = st.selectbox("Company", list(cur_sc.index))
        d = score[score.company == focus].melt(id_vars="fiscal_year", value_vars=pillars, var_name="Pillar", value_name="Score")
        fig = px.line(d, x="fiscal_year", y="Score", color="Pillar", markers=True,
                      color_discrete_sequence=[NAVY, ORANGE, "#5DA9E9", "#8C98A4", "#2E8B80", "#E4B363"],
                      labels={"fiscal_year": "Financial year"})
        comp = score[score.company == focus]
        fig.add_scatter(x=comp.fiscal_year, y=comp.Composite, name="Composite", mode="lines",
                        line=dict(color="#1E2328", width=4))
        fig.update_xaxes(dtick=1)
        fig.update_yaxes(range=[0, 100])
        st.plotly_chart(style_fig(fig, 380), width="stretch")

    with st.expander("How the score is built"):
        rows = []
        for p, items in L.PILLARS.items():
            for v, d_ in items:
                rows.append((p, L.LABELS[v], "Higher is better" if d_ > 0 else "Lower is better",
                             f"{L.CAPS[v][0]} to {L.CAPS[v][1]}" if v in L.CAPS else "none"))
        rows.append(("Stability", "Standard deviation of net margin, FY2021–FY2025", "Lower is better", "none"))
        rows.append(("ESG", "Average of scaled ESG indicators (see ESG page)", "Per indicator", "none"))
        st.dataframe(pd.DataFrame(rows, columns=["Pillar", "Measure", "Direction", "Winsorised range"]),
                     hide_index=True, width="stretch")
        st.markdown(
            "Each measure is capped at the range shown (so one extreme year, such as Valterra's 483x interest cover in "
            "FY2021, does not flatten everyone else), then min-max scaled to 0–100 over all 25 company-years. Pillar "
            "score = average of its measures. Composite = weighted average of available pillars. Bands: below 40 high "
            "risk, 40–65 moderate, above 65 low. The thresholds are a design choice and should be justified in the report."
        )


def page_esg():
    st.title("ESG indicators")
    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown(
            '<p class="lede">The workbooks contain financial statements only. Capture the sustainability indicators from '
            "each company's integrated and ESG reports into the template, with the page number for every value, "
            "then upload it here. Blank values are ignored.</p>", unsafe_allow_html=True)
    with c2:
        st.download_button("Download ESG template (CSV)", esg_template().to_csv(index=False).encode(),
                           "esg_template.csv", "text/csv", width="stretch")
        up = st.file_uploader("Upload completed ESG file", type="csv")
        if up is not None:
            try:
                df = L.load_esg(up)
                if st.session_state.get("esg_upload") != up.file_id:
                    st.session_state.esg_upload = up.file_id
                    st.session_state.esg_df = df
                    st.rerun()
                st.success(f"Loaded {len(df)} ESG values from {up.name}.")
                for w in df.attrs.get("warnings", []):
                    st.warning(w)
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not read the file: {e}")
        if not esg.empty and st.button("Remove ESG data"):
            st.session_state.pop("esg_df", None)
            st.session_state.esg_df = pd.DataFrame(columns=L.ESG_COLUMNS)
            st.rerun()
        st.caption("To load it automatically, save the file as data/esg_data.csv.")

    if esg.empty:
        st.info("No ESG values loaded yet. Once loaded, this page shows disclosure coverage, indicator trends, "
                "intensity per rand of revenue, and how ESG measures move with financial performance.")
        return

    e = esg[esg.company.isin(companies)]
    st.subheader("Disclosure coverage")
    st.caption("Share of template indicators with a value, per company and year. Gaps are themselves a finding about comparability.")
    total = esg_template().groupby(["company", "fiscal_year"]).size()
    have = e.groupby(["company", "fiscal_year"]).indicator.nunique()
    cov = (have / total).unstack().reindex([c for c in COMPANY_ORDER if c in companies]).fillna(0) * 100
    fig = go.Figure(go.Heatmap(z=cov.values, x=[str(c) for c in cov.columns], y=cov.index, zmin=0, zmax=100,
                               text=cov.round(0).astype(int).astype(str) + "%", texttemplate="%{text}",
                               colorscale=[[0, "#FAD7A0"], [1, "#6E9CC8"]], showscale=False, xgap=2, ygap=2))
    fig.update_yaxes(autorange="reversed")
    st.plotly_chart(style_fig(fig, 300, legend=False), width="stretch")

    st.subheader("Indicator trend")
    ind = st.selectbox("Indicator", sorted(e.indicator.unique()))
    d = e[e.indicator == ind].copy()
    unit = d.unit.iloc[0]
    intensity = st.toggle("Show per R million of revenue", value=False,
                          help="Divides by revenue so large and small producers can be compared.")
    if intensity:
        rev = wide.set_index(["company", "fiscal_year"]).revenue
        d["value"] = d.apply(lambda r: r.value / rev.get((r.company, r.fiscal_year), np.nan), axis=1)
        unit = f"{unit} per R million revenue"
    fig = px.line(d.sort_values("fiscal_year"), x="fiscal_year", y="value", color="company", markers=True,
                  color_discrete_map=COLOURS, category_orders={"company": COMPANY_ORDER},
                  labels={"fiscal_year": "Financial year", "value": f"{ind} ({unit})"})
    fig.update_xaxes(dtick=1)
    st.plotly_chart(style_fig(fig), width="stretch")
    src = d[["company", "fiscal_year", "value", "source_document", "report_page"]].sort_values(["company", "fiscal_year"])
    with st.expander("Sources for this indicator"):
        st.dataframe(src, hide_index=True, width="stretch")

    st.subheader("ESG against financial performance")
    c1, c2 = st.columns(2)
    x_ind = c1.selectbox("ESG indicator", sorted(e.indicator.unique()), key="x_ind")
    y_var = c2.selectbox("Financial measure", L.RATIOS, index=L.RATIOS.index("net_margin"),
                         format_func=lambda v: L.LABELS[v])
    m = e[e.indicator == x_ind][["company", "fiscal_year", "value"]].merge(
        wide[["company", "fiscal_year", y_var]], on=["company", "fiscal_year"]).dropna()
    if len(m) >= 3:
        scatter_with_fit(m, "value", y_var, x_ind, L.LABELS[y_var])
    else:
        st.info("At least three company-years with both values are needed.")


def scatter_with_fit(m: pd.DataFrame, x: str, y: str, xlab: str, ylab: str):
    fig = px.scatter(m, x=x, y=y, color="company", color_discrete_map=COLOURS, text="fiscal_year",
                     category_orders={"company": COMPANY_ORDER}, labels={x: xlab, y: ylab})
    fig.update_traces(textposition="top center", marker=dict(size=11))
    slope, icpt = np.polyfit(m[x], m[y], 1)
    xs = np.linspace(m[x].min(), m[x].max(), 20)
    fig.add_trace(go.Scatter(x=xs, y=slope * xs + icpt, mode="lines", name="Linear fit",
                             line=dict(color="#1E2328", dash="dot")))
    st.plotly_chart(style_fig(fig, 440), width="stretch")
    r = m[x].corr(m[y])
    rho = m[x].corr(m[y], method="spearman")
    n = len(m)
    t = r * np.sqrt((n - 2) / max(1 - r ** 2, 1e-9))
    st.caption(f"n = {n} company-years · Pearson r = {r:.2f} · Spearman ρ = {rho:.2f} · t = {t:.2f}. "
               "With a sample this small, treat this as descriptive evidence of association, not causation.")


def page_relationships():
    st.title("Relationships")
    st.markdown('<p class="lede">Correlations are pooled across all selected company-years. They show which ratios move '
                "together, which helps explain the risk score and checks that pillars are not double-counting.</p>",
                unsafe_allow_html=True)
    method = st.radio("Method", ["Spearman (rank)", "Pearson (linear)"], horizontal=True)
    ratio_list = ["net_margin", "roa", "roe", "current_ratio", "debt_to_equity", "liabilities_to_assets",
                  "interest_cover", "ocf_margin", "fcf_margin", "capex_intensity", "revenue_growth"]
    corr = wide[ratio_list].corr(method="spearman" if method.startswith("Spear") else "pearson")
    labels = [L.LABELS[v] for v in ratio_list]
    fig = go.Figure(go.Heatmap(z=corr.values, x=labels, y=labels, zmin=-1, zmax=1, text=corr.round(2).values,
                               texttemplate="%{text}", colorscale=[[0, "#E67E22"], [0.5, "#F7F8FA"], [1, NAVY]],
                               xgap=1, ygap=1))
    fig.update_yaxes(autorange="reversed")
    st.plotly_chart(style_fig(fig, 560, legend=False), width="stretch")

    st.subheader("Explore a pair")
    c1, c2 = st.columns(2)
    xv = c1.selectbox("Horizontal axis", ratio_list, index=ratio_list.index("capex_intensity"), format_func=lambda v: L.LABELS[v])
    yv = c2.selectbox("Vertical axis", ratio_list, index=ratio_list.index("fcf_margin"), format_func=lambda v: L.LABELS[v])
    m = wide[["company", "fiscal_year", xv, yv]].dropna()
    if xv != yv and len(m) >= 3:
        scatter_with_fit(m, xv, yv, L.LABELS[xv], L.LABELS[yv])


def page_profile():
    st.title("Company profile")
    c = st.selectbox("Company", companies)
    meta = L.COMPANIES[c]
    st.markdown(f'<p class="lede">{meta["commodity"]} · financial year ends {meta["fy_end"]} · reports in '
                f'{"US dollars" if meta["currency"] == "USD" else "rand"} · source workbook <code>{meta["file"]}</code></p>',
                unsafe_allow_html=True)
    d = wide_all[wide_all.company == c].set_index("fiscal_year")
    cur_ccy = d.currency.iloc[0]

    k = st.columns(4)
    r = d.loc[year]
    k[0].metric("Revenue", fmt_amount(r.revenue, cur_ccy))
    k[1].metric("Net profit", fmt_amount(r.net_profit, cur_ccy))
    k[2].metric("Free cash flow", fmt_amount(r.free_cash_flow, cur_ccy))
    k[3].metric("Net debt", fmt_amount(r.net_debt, cur_ccy), help="Negative = more cash than debt")

    st.subheader("Six capitals view")
    st.caption("Financial statement measures mapped to the <IR> Framework capitals. Capitals without a financial proxy "
               "need the ESG indicators.")
    e = esg[esg.company == c] if not esg.empty else pd.DataFrame(columns=L.ESG_COLUMNS)
    caps = [
        ("Financial", f"Equity {fmt_amount(r.total_equity, cur_ccy)}, debt/equity {fmt_metric('debt_to_equity', r.debt_to_equity)}, ROE {fmt_metric('roe', r.roe)}"),
        ("Manufactured", f"PPE {fmt_amount(r.ppe, cur_ccy)}, capex {fmt_amount(r.capex, cur_ccy)} ({fmt_metric('capex_intensity', r.capex_intensity)} of revenue)"),
        ("Social and relationship", f"Dividends paid {fmt_amount(r.dividends_paid, cur_ccy)}, tax expense {fmt_amount(r.tax_expense, cur_ccy)}"),
    ]
    for cap in ["Natural", "Human", "Intellectual"]:
        ee = e[(e.capital == cap) & (e.fiscal_year == year)]
        caps.append((cap, "; ".join(f"{a}: {v:,.1f} {u}" for a, v, u in zip(ee.indicator, ee.value, ee.unit))
                     or "No ESG indicators captured for this year"))
    st.dataframe(pd.DataFrame(caps, columns=["Capital", f"FY{year} evidence"]), hide_index=True, width="stretch")

    st.subheader("Five-year summary")
    show = ["revenue", "gross_profit", "ebit", "net_profit", "operating_cash_flow", "capex", "free_cash_flow",
            "dividends_paid", "total_assets", "total_equity", "total_debt", "cash"]
    t = d[show].T
    t.index = [L.LABELS[v] for v in show]
    t.columns = [f"FY{y}" for y in t.columns]
    st.dataframe(t.style.format("{:,.0f}", na_rep="–"), width="stretch")
    st.caption(f"Million, {'R' if cur_ccy != 'USD' else 'US$'}.")

    fig = go.Figure()
    fig.add_bar(x=d.index, y=d.operating_cash_flow, name="Operating cash flow", marker_color=NAVY)
    fig.add_bar(x=d.index, y=-d.capex, name="Capex", marker_color="#8C98A4")
    fig.add_bar(x=d.index, y=-d.dividends_paid, name="Dividends", marker_color=ORANGE)
    fig.add_scatter(x=d.index, y=d.net_profit, name="Net profit", mode="lines+markers", line=dict(color="#5DA9E9", width=3))
    fig.update_layout(barmode="relative", title="Where the cash went")
    fig.update_xaxes(dtick=1)
    st.plotly_chart(style_fig(fig, 420), width="stretch")


def page_data():
    st.title("Data and quality")
    st.markdown('<p class="lede">Every number in the dashboard traces back to a workbook, sheet, line item and Excel row. '
                "The brief also requires the page in the integrated report: fill in <code>report_page</code> before "
                "submission.</p>", unsafe_allow_html=True)

    qc = L.quality_checks(wide_all, long_df)
    qc = qc[qc.company.isin(companies)]
    st.subheader("Quality checks")
    counts = qc.severity.value_counts()
    k = st.columns(4)
    for col, s in zip(k, ["High", "Medium", "Low", "Info"]):
        col.metric(s, int(counts.get(s, 0)))
    st.dataframe(qc.assign(fiscal_year=qc.fiscal_year.astype("Int64")), hide_index=True, width="stretch")

    st.subheader("Extracted dataset")
    view = long_df[long_df.company.isin(companies)].copy()
    view["variable"] = view.variable.map(L.LABELS)
    st.dataframe(view, hide_index=True, width="stretch", height=380)

    st.subheader("Data dictionary")
    st.dataframe(L.DICT_DF, hide_index=True, width="stretch")

    st.subheader("Downloads")
    c1, c2, c3 = st.columns(3)
    c1.download_button("Extracted dataset (CSV)", long_df.to_csv(index=False).encode(), "dataset_extracted.csv",
                       "text/csv", width="stretch")
    c2.download_button("Analysis panel with ratios (CSV)", wide_all.round(4).to_csv(index=False).encode(),
                       "dataset_panel_ratios.csv", "text/csv", width="stretch")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        long_df.to_excel(xw, sheet_name="Extracted", index=False)
        wide_all.round(4).to_excel(xw, sheet_name="Panel and ratios", index=False)
        score_all.round(2).to_excel(xw, sheet_name="Risk scores", index=False)
        L.DICT_DF.to_excel(xw, sheet_name="Data dictionary", index=False)
        L.quality_checks(wide_all, long_df).to_excel(xw, sheet_name="Quality checks", index=False)
        if not esg.empty:
            esg.to_excel(xw, sheet_name="ESG", index=False)
    c3.download_button("Everything (Excel)", buf.getvalue(), "apfa802_dataset.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")


def page_test():
    st.title("Test the dashboard")
    st.markdown('<p class="lede">Phase 6 asks for evidence that the tool improves the use of reported information. '
                "Ask each tester to answer the five questions first from the integrated reports, then with the "
                "dashboard, and record the result here.</p>", unsafe_allow_html=True)
    with st.expander("Test tasks to give testers", expanded=True):
        st.markdown(
            "1. Which company had the weakest liquidity in FY2025?\n"
            "2. Which company's financial risk increased most between FY2021 and FY2025?\n"
            "3. Is any company paying out more in dividends than its free cash flow can support?\n"
            "4. Which company's ESG data has the weakest external assurance (provider, level and scope)?\n"
            "5. Can Scope 1 and 2 emissions be compared directly across all five companies? If not, why not?"
        )
    with st.form("feedback", clear_on_submit=True):
        c1, c2 = st.columns(2)
        role = c1.selectbox("Tester role", ["Student (accounting)", "Student (ICT)", "Lecturer", "Industry / investor", "Other"])
        method = c2.radio("Answered using", ["Integrated reports only", "Dashboard"], horizontal=True)
        c3, c4 = st.columns(2)
        correct = c3.slider("Tasks answered correctly", 0, 5, 5)
        minutes = c4.number_input("Minutes taken for all five tasks", 0.0, 240.0, 10.0, step=0.5)
        useful = st.slider("How useful was the information for making a decision? (1 = not, 5 = very)", 1, 5, 4)
        comment = st.text_area("What was missing or confusing?")
        if st.form_submit_button("Save response"):
            row = pd.DataFrame([dict(timestamp=datetime.now().isoformat(timespec="seconds"), role=role, method=method,
                                     correct=correct, minutes=minutes, usefulness=useful, comment=comment)])
            row.to_csv(FEEDBACK_FILE, mode="a", header=not FEEDBACK_FILE.exists(), index=False)
            st.success("Response saved.")

    if FEEDBACK_FILE.exists():
        fb = pd.read_csv(FEEDBACK_FILE)
        st.subheader(f"Results so far (n = {len(fb)})")
        summ = fb.groupby("method").agg(responses=("correct", "size"), avg_correct=("correct", "mean"),
                                        avg_minutes=("minutes", "mean"), avg_usefulness=("usefulness", "mean")).round(2)
        st.dataframe(summ, width="stretch")
        st.download_button("Download responses (CSV)", fb.to_csv(index=False).encode(), "validation_responses.csv", "text/csv")
        with st.expander("All responses"):
            st.dataframe(fb, hide_index=True, width="stretch")


# --------------------------------------------------------------------------------------
# Advanced analytics
# --------------------------------------------------------------------------------------
KEYS = ["company", "fiscal_year"]
S1, S2, EN, WA = "Scope 1 GHG emissions", "Scope 2 GHG emissions", "Energy consumption", "Water withdrawal"
LT, FA, WO, EMP = "Lost-time injury frequency rate (LTIFR)", "Fatalities", "Women in workforce", "Employees"

# variables that are not in the financial data dictionary: label, format, better when
ADV = {
    "Composite": ("Risk score (100 = lowest risk)", "{:.0f}", "higher"),
    "ESG": ("ESG performance score (0–100)", "{:.0f}", "higher"),
    "ESGRQ": ("Reporting quality, ESGRQ (0–100)", "{:.0f}", "higher"),
    "ghg_intensity": ("GHG intensity (t CO2e per R million revenue)", "{:,.1f}", "lower"),
    "energy_intensity": ("Energy intensity (GJ per R million revenue)", "{:,.0f}", "lower"),
    "ghg_change": ("Scope 1+2 change since first reported year", "{:+.1f}%", "lower"),
    LT: ("LTIFR (per million hours)", "{:.2f}", "lower"),
    FA: ("Fatalities", "{:.0f}", "lower"),
    WO: ("Women in workforce", "{:.1f}%", "higher"),
}
GROUP_NAMES = {2: ["Stronger profile", "Weaker profile"],
               3: ["Stronger profile", "Middle profile", "Weaker profile"],
               4: ["Strongest profile", "Upper-middle profile", "Lower-middle profile", "Weakest profile"]}
GROUP_COLOURS = {2: [NAVY, ORANGE], 3: [NAVY, "#8C98A4", ORANGE], 4: [NAVY, "#5DA9E9", "#8C98A4", ORANGE]}


def adv_label(v: str) -> str:
    return ADV[v][0] if v in ADV else L.LABELS.get(v, v)


def adv_fmt(v: str, x) -> str:
    if pd.isna(x):
        return "–"
    return ADV[v][1].format(x) if v in ADV else fmt_metric(v, x)


def adv_better(v: str) -> str:
    return ADV[v][2] if v in ADV else L.DICT_DF.set_index("variable").better_when.get(v, "higher")


@st.cache_data(show_spinner="Reading the report...")
def scan_pdf(raw: bytes) -> pd.DataFrame:
    return NLP.theme_counts(NLP.extract_pages(raw))


def adv_panel() -> pd.DataFrame:
    """One row per company-year: ratios, risk score, ESG indicators and the ESG ratios derived from them."""
    p = wide.merge(score[KEYS + ["Composite"] + (["ESG"] if "ESG" in score else [])], on=KEYS, how="left")
    if "ESG" not in p:
        p["ESG"] = np.nan
    if not esg.empty:
        e = esg.pivot_table(index=KEYS, columns="indicator", values="value", aggfunc="first").reset_index()
        e.columns.name = None
        p = p.merge(e, on=KEYS, how="left")
    for c in (S1, S2, EN, WA, LT, FA, WO, EMP):
        if c not in p:
            p[c] = np.nan
    # intensities always use rand revenue so Gold Fields is comparable whatever the sidebar toggle says
    rand = get_wide(tuple(sorted(fx.items())))[KEYS + ["revenue"]].rename(columns={"revenue": "revenue_zar"})
    p = p.merge(rand, on=KEYS, how="left")
    p["ghg_total"] = p[S1] + p[S2]
    p["ghg_intensity"] = p.ghg_total / p.revenue_zar
    p["energy_intensity"] = p[EN] / p.revenue_zar
    first = p.dropna(subset=["ghg_total"]).sort_values("fiscal_year").groupby("company").ghg_total.first()
    p["ghg_change"] = (p.ghg_total / p.company.map(first) - 1) * 100
    return p.sort_values(KEYS).reset_index(drop=True)


def esgrq_latest():
    """(workbook sheets, report year, composite table indexed by company) or (None, None, None)."""
    data = get_esgrq()
    if data is None or data["Coding"].report_year.dropna().empty:
        return None, None, None
    ry = int(data["Coding"].report_year.dropna().max())
    comp = Q.composite(data["Coding"][data["Coding"].report_year == ry])
    return data, ry, comp.set_index("company")


def adv_anomalies(p: pd.DataFrame, cut: float) -> pd.DataFrame:
    rows = []

    def add(r, metric, value, signal, kind):
        rows.append(dict(Company=r.company, Year=int(r.fiscal_year), Metric=metric, Value=value, Signal=signal, Type=kind))

    # 1. robust z-score: distance from the sample median in units of the median absolute deviation
    for v in ["net_margin", "roe", "current_ratio", "debt_to_equity", "interest_cover", "fcf_margin",
              "revenue_growth", "ghg_intensity", "energy_intensity", LT]:
        s = p[v].dropna()
        if len(s) < 5:
            continue
        med = s.median()
        mad = (s - med).abs().median()
        if not mad:
            continue
        z = 0.6745 * (p[v] - med) / mad
        for i in z[z.abs() >= cut].index:
            add(p.loc[i], adv_label(v), adv_fmt(v, p.at[i, v]),
                f"Robust z = {z[i]:+.1f}; sample median {adv_fmt(v, med)}", "Statistical outlier")

    # 2. business rules
    for _, r in p.iterrows():
        if r.net_profit < 0:
            add(r, "Net profit", f"{r.net_profit:,.0f} million", "Loss for the year", "Business rule")
        if pd.notna(r.interest_cover) and r.interest_cover < 1.5:
            add(r, L.LABELS["interest_cover"], fmt_metric("interest_cover", r.interest_cover),
                "Earnings cover interest less than 1.5 times", "Business rule")
        if pd.notna(r.current_ratio) and r.current_ratio < 1:
            add(r, L.LABELS["current_ratio"], fmt_metric("current_ratio", r.current_ratio),
                "Current liabilities exceed current assets", "Business rule")
        if pd.notna(r.free_cash_flow) and r.dividends_paid > 0 and r.dividends_paid > r.free_cash_flow:
            add(r, "Dividends against free cash flow", f"{r.dividends_paid:,.0f} vs {r.free_cash_flow:,.0f} million",
                "Dividends paid exceed free cash flow", "Business rule")

    # 3. ESG values that jump between years: often a boundary change or a capture error
    for ind in (S1, S2, EN, WA, EMP):
        chg = p.groupby("company")[ind].pct_change(fill_method=None) * 100
        for i in chg[chg.abs() >= 40].index:
            add(p.loc[i], ind, f"{p.at[i, ind]:,.0f}", f"{chg[i]:+.0f}% on the previous year: check boundary and capture",
                "ESG movement")

    # 4. capture problems already found by the data-quality checks
    qc = L.quality_checks(wide_all, long_df)
    for r in qc[(qc.severity == "High") & qc.company.isin(companies)].itertuples():
        rows.append(dict(Company=r.company, Year=int(r.fiscal_year) if pd.notna(r.fiscal_year) else None,
                         Metric=r.check, Value="", Signal=r.detail, Type="Data capture check"))

    out = pd.DataFrame(rows, columns=["Company", "Year", "Metric", "Value", "Signal", "Type"])
    return out.sort_values(["Company", "Year", "Type"]).reset_index(drop=True)


def adv_insights(p: pd.DataFrame, comp, data, ry, anomalies: pd.DataFrame, themes: pd.DataFrame) -> list[tuple[str, bool]]:
    out = []
    order = [c for c in COMPANY_ORDER if c in companies]
    if comp is not None:
        q = comp.reindex(order).ESGRQ.dropna()
        if len(q) >= 2:
            out.append((f"<b>{q.idxmax()}</b> has the highest ESG reporting quality in the {ry} reports "
                        f"(ESGRQ {q.max():.0f}/100) and <b>{q.idxmin()}</b> the lowest ({q.min():.0f}/100).", False))
        _, cscore = Q.comparability(data["Comparability"][data["Comparability"].report_year == ry], order)
        if cscore is not None:
            share = "None" if cscore < 0.5 else f"Only <b>{cscore:.0f}%</b>" if cscore < 50 else f"<b>{cscore:.0f}%</b>"
            out.append((f"{share} of the common ESG indicators meet every comparability criterion (same unit, "
                        "boundary and period), so ESG figures cannot be compared across the companies without "
                        "adjustment." if cscore < 50 else
                        f"{share} of the common ESG indicators meet every comparability criterion (same unit, "
                        "boundary and period).", cscore < 50))
        lv = data["Assurance"][data["Assurance"].company.isin(order)]
        limited = lv[lv.level.fillna("").str.lower().str.startswith("limited")].company.tolist()
        if limited:
            out.append((f"<b>{', '.join(limited)}</b> obtained limited assurance only over ESG data, the weakest level "
                        "in the sample. Treat those disclosures with more caution.", True))
    m = p[["ESG", "roe"]].dropna()
    if len(m) >= 5:
        r = m.ESG.corr(m.roe)
        word = "positive" if r > 0.3 else "negative" if r < -0.3 else "weak"
        out.append((f"ESG performance and return on equity show a <b>{word}</b> association across {len(m)} "
                    f"company-years (r = {r:.2f}). This is descriptive: the sample is too small to claim cause.", False))
    if not anomalies.empty:
        top = anomalies.Company.value_counts()
        out.append((f"<b>{len(anomalies)}</b> observations are flagged as unusual; <b>{top.index[0]}</b> has the most "
                    f"({top.iloc[0]}). See the Anomaly detection tab for each one.", True))
    if not esg.empty:
        n = esg[esg.company.isin(order)].groupby("company").size().reindex(order).fillna(0)
        if n.min() < n.max() * 0.5:
            out.append((f"<b>{n.idxmin()}</b> has only {n.min():.0f} ESG indicator values captured, against "
                        f"{n.max():.0f} for {n.idxmax()}. Its ESG charts and ESG score rest on thin data.", True))
    if not themes.empty:
        t = themes[themes.company.isin(order)].pivot_table(index="company", columns="theme", values="mentions")
        if {"Commitment language", "Performance language"} <= set(t.columns):
            ratio = (t["Commitment language"] / t["Performance language"]).dropna()
            if len(ratio) >= 2:
                out.append((f"<b>{ratio.idxmax()}</b>'s report uses the most forward-looking language relative to "
                            f"performance language ({ratio.max():.1f} commitment words per performance word; "
                            f"{ratio.idxmin()} is lowest at {ratio.min():.1f}).", False))
    return out


def trace_table(data, themes: pd.DataFrame) -> pd.DataFrame:
    """Every value the dashboard uses, with the document and the place in it that the value came from."""
    cols = ["Dataset", "Company", "Year", "Item", "Value", "Unit", "Source document", "Location in source"]
    parts = []
    f = long_df[long_df.company.isin(companies)]
    page = f.report_page.fillna("").astype(str).str.strip()
    parts.append(pd.DataFrame({
        "Dataset": "Financial statement line", "Company": f.company, "Year": f.fiscal_year,
        "Item": f.variable.map(L.LABELS), "Value": f.value_m.map(lambda v: f"{v:,.1f}"),
        "Unit": f.currency + " million", "Source document": f.source_file,
        "Location in source": f.source_lines + page.map(lambda s: f"; report p.{s}" if s else ""),
    }))
    if not esg.empty:
        e = esg[esg.company.isin(companies)]
        parts.append(pd.DataFrame({
            "Dataset": "ESG indicator", "Company": e.company, "Year": e.fiscal_year, "Item": e.indicator,
            "Value": e.value.map(lambda v: f"{v:,.2f}".rstrip("0").rstrip(".")), "Unit": e.unit,
            "Source document": e.source_document, "Location in source": e.report_page,
        }))
    if data is not None:
        c = data["Coding"]
        c = c[c.company.isin(companies) & c.score.notna()]
        parts.append(pd.DataFrame({
            "Dataset": "Reporting quality score", "Company": c.company, "Year": c.report_year,
            "Item": c.code + " " + c.indicator, "Value": c.score.map(lambda v: f"{v:.0f}"), "Unit": "score 0–4",
            "Source document": c.source_document, "Location in source": c.report_page.map(lambda s: f"p.{s}" if s else ""),
        }))
    if not themes.empty:
        t = themes[themes.company.isin(companies)]
        parts.append(pd.DataFrame({
            "Dataset": "Report text theme", "Company": t.company, "Year": 2025, "Item": t.theme,
            "Value": t.mentions.map(lambda v: f"{v:,.0f}"), "Unit": "mentions", "Source document": t.source_document,
            "Location in source": t.example_page.map(lambda v: f"example on PDF p.{v:.0f}" if pd.notna(v) else ""),
        }))
    out = pd.concat(parts, ignore_index=True)[cols]
    for c in ("Source document", "Location in source"):
        out[c] = out[c].fillna("").astype(str).str.strip()
    return out


def full_dictionary() -> pd.DataFrame:
    cols = ["Variable", "Type", "Unit / measurement", "Definition", "Better when"]
    fin = pd.DataFrame({"Variable": L.DICT_DF.label, "Type": "Financial: " + L.DICT_DF.type.str.lower(),
                        "Unit / measurement": L.DICT_DF.unit, "Definition": L.DICT_DF.definition,
                        "Better when": L.DICT_DF.better_when})
    t = esg_template().drop_duplicates("indicator")
    pillar = {"E": "Environmental", "S": "Social", "G": "Governance"}
    ind = pd.DataFrame({"Variable": t.indicator, "Type": "ESG indicator (" + t.pillar.map(pillar) + ")",
                        "Unit / measurement": t.unit,
                        "Definition": "Captured from the company's ESG or sustainability report; " + t.capital + " capital.",
                        "Better when": t.better_when})
    q = pd.DataFrame({"Variable": Q.ITEM_TABLE.code + " " + Q.ITEM_TABLE.indicator,
                      "Type": "Reporting quality item (" + Q.ITEM_TABLE.dimension + ")",
                      "Unit / measurement": "Score 0–4",
                      "Definition": "Evidence required: " + Q.ITEM_TABLE.evidence_required, "Better when": "higher"})
    derived = pd.DataFrame([
        ("Risk score", "Constructed index", "0–100", "Weighted mean of the pillar scores (profitability, liquidity, solvency, cash generation, stability and, if weighted, ESG). 100 = lowest risk in the sample.", "higher"),
        ("ESG performance score", "Constructed index", "0–100", "Average of the ESG indicators after scaling each one 0–100 across the sample, direction taken into account.", "higher"),
        ("ESGRQ", "Constructed index", "0–100", "ESG Reporting Quality: equal-weighted mean of the six dimension scores, each = mean item score ÷ 4 × 100.", "higher"),
        ("GHG intensity", "Calculated ratio", "t CO2e per R million revenue", "(Scope 1 + Scope 2 emissions) ÷ revenue in rand.", "lower"),
        ("Energy intensity", "Calculated ratio", "GJ per R million revenue", "Energy consumption ÷ revenue in rand.", "lower"),
        ("Scope 1+2 change since first reported year", "Calculated", "%", "Scope 1 + 2 emissions against the earliest year captured for that company.", "lower"),
        ("Anomaly flag", "Analytical flag", "Robust z-score / rule", "Value at least the chosen number of robust z-scores (0.6745 × distance from median ÷ median absolute deviation) from the sample median, or one that breaks a business rule.", "n/a"),
        ("Benchmark position", "Analytical flag", "Above / below peer average", "Company value against the mean of the selected peers for the selected year, direction taken into account.", "n/a"),
        ("Regression slope, intercept, R²", "Statistic", "Numeric", "Ordinary least squares fit of one variable on another across company-years; p-value from a two-sided t-test.", "n/a"),
        ("Profile group", "Analytical grouping", "Category", "Group from Ward hierarchical clustering of standardised net margin, ROE, current ratio, debt/equity and FCF margin.", "n/a"),
        ("Report text theme", "Text analytics", "Mentions per 100 pages", "Count of theme keywords in the report text, with one example sentence and its PDF page.", "n/a"),
    ], columns=cols)
    return pd.concat([fin, ind, q, derived], ignore_index=True)[cols]


def page_advanced():
    st.title("Advanced analytics")
    st.markdown(
        '<p class="lede">Multi-year ratios, anomaly detection, benchmarking, comparison, regression, report-text '
        "analysis, clustering, scoring and descriptive statistics on the group's captured data, followed by the "
        "insight summary, the traceable dataset and the data dictionary.</p>", unsafe_allow_html=True)

    p = adv_panel()
    order = [c for c in COMPANY_ORDER if c in companies]
    data, ry, comp = esgrq_latest()
    themes = NLP.load()
    now = p[p.fiscal_year == year].set_index("company").reindex(order)
    now["ESGRQ"] = comp.ESGRQ.reindex(order) if comp is not None else np.nan

    tabs = st.tabs(["Ratios & years", "Anomaly detection", "Benchmarking", "Comparative analysis",
                    "Correlation & regression", "Natural language processing", "Clustering", "Scoring index",
                    "Descriptive analysis", "Data visualisation"])

    # ---- Ratios & years --------------------------------------------------------------
    with tabs[0]:
        k = st.columns(4)
        yrs = sorted(p.fiscal_year.unique())
        k[0].metric("Years covered", len(yrs), f"FY{yrs[0]}–FY{yrs[-1]}", delta_color="off")
        k[1].metric("Companies", len(order))
        k[2].metric("Company-years", len(p))
        k[3].metric("ESG values captured", 0 if esg.empty else int(esg.company.isin(companies).sum()))
        show = ["net_margin", "roe", "revenue_growth", "current_ratio", "debt_to_equity", "fcf_margin",
                "ghg_intensity", "ghg_change", LT]
        t = p[KEYS].rename(columns={"company": "Company", "fiscal_year": "Year"})
        for v in show:
            t[adv_label(v)] = p[v].map(lambda x, v=v: adv_fmt(v, x))
        st.dataframe(t, hide_index=True, width="stretch", height=420)
        st.caption("Financial ratios come from the financial-statement workbooks. GHG intensity and LTIFR come from "
                   "the ESG indicators; a dash means the value has not been captured for that year.")

    # ---- Anomaly detection -----------------------------------------------------------
    with tabs[1]:
        c1, c2 = st.columns([3, 1])
        cut = c1.slider("Sensitivity: flag values this many robust z-scores from the sample median", 2.5, 5.0, 3.5, 0.5,
                        help="3.5 is the usual cut-off for the median-based z-score. Lower flags more observations.")
        anomalies = adv_anomalies(p, cut)
        c2.metric("Flagged observations", len(anomalies))
        if anomalies.empty:
            st.success("Nothing unusual at this sensitivity.")
        else:
            st.dataframe(anomalies, hide_index=True, width="stretch", height=420,
                         column_config={"Year": st.column_config.NumberColumn(format="%d")})
            by = anomalies.groupby(["Company", "Type"]).size().reset_index(name="Flags")
            fig = px.bar(by, x="Company", y="Flags", color="Type", category_orders={"Company": order},
                         color_discrete_sequence=[NAVY, ORANGE, "#5DA9E9", "#8C98A4"], labels={"Company": ""})
            st.plotly_chart(style_fig(fig, 320), width="stretch")
        st.caption("A flag is a prompt to look at the report, not proof of an error. Statistical outliers use the "
                   "median and median absolute deviation, which a few extreme years cannot distort.")

    # ---- Benchmarking ----------------------------------------------------------------
    with tabs[2]:
        bench = [v for v in ["Composite", "net_margin", "roe", "current_ratio", "debt_to_equity", "fcf_margin",
                             "ESGRQ", "ghg_intensity", LT] if now[v].notna().sum() >= 2]
        if not bench:
            st.info("Benchmarking needs at least two selected companies with data for the year.")
        else:
            avg = now[bench].mean()
            t = pd.DataFrame(index=order)
            above = pd.Series(0, index=order)
            for v in bench:
                t[adv_label(v)] = now[v].map(lambda x, v=v: adv_fmt(v, x))
                ahead = now[v] < avg[v] if adv_better(v) == "lower" else now[v] > avg[v]
                above += ahead.fillna(False).astype(int)
            counted = now[bench].notna().sum(axis=1)
            t["Better than peer average"] = [f"{a} of {n}" for a, n in zip(above, counted)]
            t.loc["Peer average"] = [adv_fmt(v, avg[v]) for v in bench] + [""]
            t.index.name = "Company"
            st.dataframe(t, width="stretch")
            var = st.selectbox("Chart a measure against the peer average", bench, format_func=adv_label, key="adv_bench")
            d = now[[var]].dropna().reset_index()
            fig = px.bar(d, x="company", y=var, color="company", color_discrete_map=COLOURS,
                         text=d[var].map(lambda x: adv_fmt(var, x)), labels={"company": "", var: adv_label(var)})
            fig.add_hline(y=avg[var], line_dash="dot", line_color="#1E2328",
                          annotation_text=f"Peer average {adv_fmt(var, avg[var])}", annotation_position="top left")
            st.plotly_chart(style_fig(fig, 360, legend=False), width="stretch")
            st.caption(f"FY{year}, selected companies. For {adv_label(var).lower()}, {adv_better(var)} is better. "
                       "ESGRQ is scored on the 2025 reports whichever year is selected.")

    # ---- Comparative analysis --------------------------------------------------------
    with tabs[3]:
        t = pd.DataFrame(index=order)
        t["Risk band"] = cur_sc["Risk band"].reindex(order).astype(str).replace("nan", "–")
        for v in ["Composite", "ESGRQ", "roe", "net_margin", "debt_to_equity", "ghg_intensity", LT, FA, WO]:
            t[adv_label(v)] = now[v].map(lambda x, v=v: adv_fmt(v, x))
        if data is not None:
            a = data["Assurance"].drop_duplicates("company").set_index("company")
            t["ESG assurance level"] = a.level.reindex(order).fillna("–")
            t["Assurance provider"] = a.provider.reindex(order).fillna("–")
        t.index.name = "Company"
        st.dataframe(t.T, width="stretch", height=460)
        st.caption(f"FY{year} side by side: financial risk, reporting quality, ESG performance and assurance. "
                   "Assurance and ESGRQ refer to the 2025 reports.")

    # ---- Correlation & regression ----------------------------------------------------
    with tabs[4]:
        reg = [v for v in ["ESG", "Composite", "net_margin", "roe", "roa", "current_ratio", "debt_to_equity",
                           "interest_cover", "fcf_margin", "capex_intensity", "revenue_growth", "ghg_intensity",
                           "energy_intensity", LT, FA, WO] if p[v].notna().sum() >= 3]
        c1, c2 = st.columns(2)
        xv = c1.selectbox("Explanatory variable (x)", reg, index=reg.index("ESG") if "ESG" in reg else 0,
                          format_func=adv_label, key="adv_x")
        yv = c2.selectbox("Outcome variable (y)", reg, index=reg.index("roe"), format_func=adv_label, key="adv_y")
        m = p[KEYS + [xv, yv]].dropna() if xv != yv else pd.DataFrame()
        if len(m) < 3 or m[xv].nunique() < 2:
            st.info("Choose two different variables with at least three company-years in common.")
        else:
            fit = linregress(m[xv], m[yv])
            k = st.columns(5)
            k[0].metric("Correlation (r)", f"{fit.rvalue:+.3f}")
            k[1].metric("R²", f"{fit.rvalue ** 2:.3f}")
            k[2].metric("Slope", f"{fit.slope:,.3f}")
            k[3].metric("Intercept", f"{fit.intercept:,.2f}")
            k[4].metric("p-value", f"{fit.pvalue:.3f}", f"n = {len(m)}", delta_color="off")
            fig = px.scatter(m, x=xv, y=yv, color="company", color_discrete_map=COLOURS, text="fiscal_year",
                             category_orders={"company": COMPANY_ORDER}, labels={xv: adv_label(xv), yv: adv_label(yv)})
            fig.update_traces(textposition="top center", marker=dict(size=11))
            xs = np.linspace(m[xv].min(), m[xv].max(), 20)
            fig.add_trace(go.Scatter(x=xs, y=fit.intercept + fit.slope * xs, mode="lines", name="OLS regression line",
                                     line=dict(color="#1E2328", dash="dot")))
            st.plotly_chart(style_fig(fig, 440), width="stretch")
            sig = "statistically significant at the 5% level" if fit.pvalue < 0.05 else "not statistically significant at the 5% level"
            st.markdown(f'<div class="insight"><b>Regression equation:</b> {adv_label(yv)} = {fit.intercept:,.2f} + '
                        f"({fit.slope:,.3f} × {adv_label(xv)}). The fit explains {fit.rvalue ** 2 * 100:.0f}% of the "
                        f"variation and is {sig}.</div>", unsafe_allow_html=True)
            st.caption("Pooled ordinary least squares across company-years. Observations from the same company are not "
                       "independent and the sample is small, so read this as association, not causation.")

    # ---- Natural language processing -------------------------------------------------
    with tabs[5]:
        th = themes[themes.company.isin(companies)] if not themes.empty else themes
        if th.empty:
            st.info("No report text has been scanned yet. Run `python nlp.py` with the report PDFs in the folder above "
                    "the app, or upload a report below.")
        else:
            st.caption("The five 2025 ESG and sustainability reports were read page by page. Each theme is a set of "
                       "keywords; the number is mentions per 100 pages so long and short reports can be compared.")
            topic = th[th.group == "Topic"].pivot_table(index="theme", columns="company", values="per_100_pages")
            topic = topic.reindex(index=[t for t, (g, _) in NLP.THEMES.items() if g == "Topic"], columns=order)
            st.plotly_chart(heat(topic, zmax=float(np.nanmax(topic.values)), height=440,
                                 scale=[[0, "#F4F7FA"], [1, NAVY]]), width="stretch")

            st.subheader("Commitment, performance and risk language")
            lang = th[th.group == "Language"]
            fig = px.bar(lang, x="company", y="per_10k_words", color="theme", barmode="group",
                         category_orders={"company": order}, color_discrete_sequence=[ORANGE, NAVY, "#8C98A4"],
                         labels={"company": "", "per_10k_words": "Mentions per 10,000 words"})
            st.plotly_chart(style_fig(fig, 340), width="stretch")
            piv = lang.pivot_table(index="company", columns="theme", values="mentions").reindex(order)
            ratio = piv["Commitment language"] / piv["Performance language"]
            st.caption("Commitment words per performance word: "
                       + " · ".join(f"{c} {v:.1f}" for c, v in ratio.dropna().items())
                       + ". A report that promises far more than it reports on is worth a closer look at whether the "
                         "targets are backed by numbers.")

            st.subheader("Evidence")
            pick = st.selectbox("Theme", list(NLP.THEMES), key="adv_theme")
            ev = th[th.theme == pick].set_index("company").reindex(order).reset_index()
            st.dataframe(pd.DataFrame({
                "Company": ev.company, "Mentions": ev.mentions, "Pages mentioning it": ev.pages_mentioning,
                "Report pages": ev.report_pages, "Example page (PDF)": ev.example_page, "Example": ev.example,
                "Source": ev.source_document}), hide_index=True, width="stretch")

        with st.expander("Scan another report (PDF)"):
            up = st.file_uploader("Report PDF with selectable text", type="pdf", key="adv_pdf")
            if up is not None:
                try:
                    res = scan_pdf(up.getvalue())
                    st.success(f"Read {res.report_pages.iloc[0]} pages and {res.report_words.iloc[0]:,} words from {up.name}.")
                    st.dataframe(res[["theme", "group", "mentions", "pages_mentioning", "per_100_pages",
                                      "example_page", "example"]], hide_index=True, width="stretch")
                except Exception as e:  # noqa: BLE001
                    st.error(f"Could not read this PDF: {e}. Scanned, image-only reports need OCR first.")
        st.caption("Keyword counts show what a report talks about, not how well the company performs. They support "
                   "the coded reporting-quality scores; they do not replace them.")

    # ---- Clustering ------------------------------------------------------------------
    with tabs[6]:
        feats = ["net_margin", "roe", "current_ratio", "debt_to_equity", "fcf_margin"]
        d = p.dropna(subset=feats).copy()
        if len(d) < 6:
            st.info("Clustering needs at least six company-years with complete ratios.")
        else:
            kk = st.radio("Number of groups", [2, 3, 4], index=1, horizontal=True, key="adv_k")
            X = d[feats].copy()
            for v in feats:  # limit extreme years so one loss-making year does not define the groups
                X[v] = X[v].clip(*L.CAPS[v]) if v in L.CAPS else X[v].clip(X[v].quantile(0.05), X[v].quantile(0.95))
            Z = (X - X.mean()) / X.std(ddof=0).replace(0, 1)
            d["g"] = fcluster(linkage(Z.values, "ward"), kk, "maxclust")
            rank = d.groupby("g").Composite.mean().sort_values(ascending=False).index.tolist()
            names = GROUP_NAMES[kk][:len(rank)]
            d["Group"] = d.g.map(dict(zip(rank, names)))
            U, S, _ = np.linalg.svd(Z.values, full_matrices=False)
            d["pc1"], d["pc2"] = U[:, 0] * S[0], U[:, 1] * S[1]
            share = S ** 2 / (S ** 2).sum()
            d["label"] = d.company.map(lambda c: L.COMPANIES[c]["ticker"]) + " " + d.fiscal_year.astype(str).str[-2:]

            cols = st.columns(len(names))
            for col, name in zip(cols, names):
                g = d[d.Group == name]
                members = "<br>".join(f"<b>{c}</b>: " + ", ".join(f"FY{y}" for y in sorted(gg.fiscal_year))
                                      for c, gg in g.groupby("company", sort=False))
                col.markdown(f'<div class="insight"><b>{name}</b> · {len(g)} company-years<br>'
                             f'<small>{members}</small></div>', unsafe_allow_html=True)
            fig = px.scatter(d, x="pc1", y="pc2", color="Group", text="label", category_orders={"Group": names},
                             color_discrete_sequence=GROUP_COLOURS[kk],
                             labels={"pc1": f"Component 1 ({share[0] * 100:.0f}% of variation)",
                                     "pc2": f"Component 2 ({share[1] * 100:.0f}% of variation)"})
            fig.update_traces(textposition="top center", marker=dict(size=12))
            st.plotly_chart(style_fig(fig, 460), width="stretch")
            means = d.groupby("Group")[feats + ["Composite"]].mean().reindex(names)
            st.dataframe(pd.DataFrame({adv_label(v): means[v].map(lambda x, v=v: adv_fmt(v, x)) for v in feats + ["Composite"]}),
                         width="stretch")
            st.caption("Ward hierarchical clustering on standardised net margin, ROE, current ratio, debt/equity and FCF "
                       "margin. Each point is one company-year (ticker and year); the two axes are the main directions "
                       "of difference between them. Groups are named by their average risk score. ESG indicators are "
                       "left out because too many company-years have gaps.")

    # ---- Scoring index ---------------------------------------------------------------
    with tabs[7]:
        pillars = [x for x in ["Profitability", "Liquidity", "Solvency", "Cash generation", "Stability", "ESG"] if x in cur_sc]
        t = cur_sc[pillars + ["Composite"]].reindex(order).rename(columns={"Composite": "Risk score"})
        t["ESGRQ (2025 reports)"] = now.ESGRQ
        t = t.sort_values("Risk score", ascending=False)
        st.dataframe(t.style.format("{:.0f}", na_rep="–"), width="stretch")
        long = t[["Risk score", "ESGRQ (2025 reports)"]].reset_index().melt("company", var_name="Index", value_name="Score")
        fig = px.bar(long, x="company", y="Score", color="Index", barmode="group", text=long.Score.map(lambda v: "" if pd.isna(v) else f"{v:.0f}"),
                     color_discrete_sequence=[NAVY, ORANGE], labels={"company": "", "Score": "Index (0–100)"})
        fig.update_yaxes(range=[0, 105])
        st.plotly_chart(style_fig(fig, 380), width="stretch")
        w = st.session_state.weights
        st.caption(f"FY{year}. The risk score is a weighted mean of the pillar scores ("
                   + ", ".join(f"{k} {v}%" for k, v in w.items() if v) + "); change the weights on the Risk scorecard "
                   "page. ESGRQ measures the quality of ESG reporting, not ESG performance. Neither index is an "
                   "investment recommendation.")

    # ---- Descriptive analysis --------------------------------------------------------
    with tabs[8]:
        dv = ["revenue_growth", "net_margin", "roe", "roa", "current_ratio", "debt_to_equity", "interest_cover",
              "fcf_margin", "Composite", "ESG", "ghg_intensity", "energy_intensity", LT, FA, WO]
        rows = []
        for v in dv:
            s = p[v].dropna()
            if s.empty:
                continue
            rows.append({"Variable": adv_label(v), "n": len(s), "Mean": adv_fmt(v, s.mean()), "Median": adv_fmt(v, s.median()),
                         "Minimum": adv_fmt(v, s.min()), "Maximum": adv_fmt(v, s.max()),
                         "Std. dev.": adv_fmt(v, s.std()) if len(s) > 1 else "–"})
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch", height=38 * (len(rows) + 1))
        st.caption("Pooled over the selected companies and FY2021–FY2025. n is the number of company-years with a "
                   "value; the standard deviation is the sample standard deviation.")
        var = st.selectbox("Distribution by company", dv, index=1, format_func=adv_label, key="adv_box")
        fig = px.box(p.dropna(subset=[var]), x="company", y=var, color="company", color_discrete_map=COLOURS,
                     points="all", category_orders={"company": order}, labels={"company": "", var: adv_label(var)})
        st.plotly_chart(style_fig(fig, 380, legend=False), width="stretch")

    # ---- Data visualisation ----------------------------------------------------------
    with tabs[9]:
        idx = [x for x in ["Profitability", "Liquidity", "Solvency", "Cash generation", "Stability", "ESG"] if x in cur_sc]
        prof = cur_sc[idx].reindex(order)
        prof["ESGRQ"] = now.ESGRQ
        prof = prof.dropna(axis=1, how="all")
        fig = go.Figure()
        for c in order:
            vals = prof.loc[c].tolist()
            fig.add_trace(go.Scatterpolar(r=vals + vals[:1], theta=list(prof.columns) + [prof.columns[0]], name=c,
                                          mode="lines", line=dict(color=COLOURS[c], width=2)))
        fig.update_layout(polar=dict(radialaxis=dict(range=[0, 100], showticklabels=False, gridcolor="#E8ECF1"),
                                     angularaxis=dict(gridcolor="#E8ECF1"), bgcolor="#FFFFFF"))
        fig = style_fig(fig, 480).update_layout(margin=dict(l=70, r=70, t=40, b=100),
                                                legend=dict(orientation="h", y=-0.2, yanchor="top", x=0.5, xanchor="center"))
        st.plotly_chart(fig, width="stretch")
        long = prof.reset_index().melt("company", var_name="Measure", value_name="Score")
        fig = px.bar(long, x="Measure", y="Score", color="company", barmode="group", color_discrete_map=COLOURS,
                     category_orders={"company": order}, labels={"Measure": "", "Score": "Indexed score (0–100)"})
        fig.update_yaxes(range=[0, 105])
        st.plotly_chart(style_fig(fig, 400), width="stretch")
        st.caption(f"FY{year}. Every measure is on the same 0–100 scale, 100 being the strongest position in the sample, "
                   "so financial strength, ESG performance and reporting quality (ESGRQ) can be read together.")

    # ---- Insight summary -------------------------------------------------------------
    st.subheader("Automated insight summary")
    st.caption(f"Generated from the current selection (FY{year}, {len(order)} companies). A starting point for the "
               "results and discussion sections; check each statement against the cited source before quoting it.")
    for text, warn in insights() + adv_insights(p, comp, data, ry, adv_anomalies(p, 3.5), themes):
        st.markdown(f'<div class="insight{" warn" if warn else ""}">{text}</div>', unsafe_allow_html=True)

    # ---- Traceable dataset -----------------------------------------------------------
    st.subheader("Traceable dataset")
    tr = trace_table(data, themes)
    sets = list(dict.fromkeys(tr.Dataset))
    pick = st.multiselect("Show", sets, default=sets, key="adv_sets")
    view = tr[tr.Dataset.isin(pick)]
    has_page = view["Location in source"].str.contains(r"p\.\s?\d|row \d|Sheet", case=False, regex=True)
    k = st.columns(3)
    k[0].metric("Observations", f"{len(view):,}")
    k[1].metric("With a source document", f"{(view['Source document'] != '').mean() * 100:.0f}%" if len(view) else "–")
    k[2].metric("With a page, sheet or row", f"{has_page.mean() * 100:.0f}%" if len(view) else "–")
    st.dataframe(view, hide_index=True, width="stretch", height=380,
                 column_config={"Year": st.column_config.NumberColumn(format="%d")})
    st.caption("Every value the dashboard uses, with the document it came from and where in that document. Financial "
               "lines trace to the workbook sheet and row; ESG values, reporting-quality scores and text themes trace "
               "to the report page.")
    st.download_button("Download traceable dataset (CSV)", tr.to_csv(index=False).encode("utf-8-sig"),
                       "traceable_dataset.csv", "text/csv")

    # ---- Data dictionary -------------------------------------------------------------
    st.subheader("Data dictionary")
    dd = full_dictionary()
    st.dataframe(dd, hide_index=True, width="stretch", height=380)
    st.download_button("Download data dictionary (CSV)", dd.to_csv(index=False).encode("utf-8-sig"),
                       "data_dictionary.csv", "text/csv")


{
    "Overview": page_overview, "Reporting quality": page_esgrq, "Trends": page_trends, "Peer benchmark": page_benchmark,
    "Risk scorecard": page_scorecard, "ESG indicators": page_esg, "Relationships": page_relationships,
    "Advanced analytics": page_advanced, "Company profile": page_profile,
    "Data and quality": page_data, "Test the dashboard": page_test,
}[page]()
