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

import esgrq as Q
import loader as L

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
         "Relationships", "Company profile", "Data and quality", "Test the dashboard"]
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


def heat(z: pd.DataFrame, text: pd.DataFrame | None = None, zmax: float = 100, height: int = 300) -> go.Figure:
    t = text if text is not None else z.round(0).map(lambda v: "" if pd.isna(v) else f"{v:.0f}")
    fig = go.Figure(go.Heatmap(z=z.values, x=list(z.columns), y=list(z.index), zmin=0, zmax=zmax, text=t.values,
                               texttemplate="%{text}", colorscale=SEQ, showscale=False, xgap=2, ygap=2,
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


{
    "Overview": page_overview, "Reporting quality": page_esgrq, "Trends": page_trends, "Peer benchmark": page_benchmark,
    "Risk scorecard": page_scorecard, "ESG indicators": page_esg, "Relationships": page_relationships,
    "Company profile": page_profile, "Data and quality": page_data, "Test the dashboard": page_test,
}[page]()
