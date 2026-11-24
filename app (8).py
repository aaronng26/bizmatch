"""
Business Matchmaking Dashboard (Streamlit)

Run:
    pip install -r requirements.txt
    streamlit run app.py

Then upload the customer data file (Excel or CSV) in the sidebar. The Excel
template business_matchmaking_template.xlsx shows the expected columns.
Every figure is computed from that file; nothing is sent anywhere.
"""

from __future__ import annotations

import io
import math
import re
from datetime import date

import difflib
import html

import altair as alt
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

# Known main sectors. Sector names found in the data are matched to these (and to each other)
# ignoring case, plurals, "and"/"&" and small typos, e.g. "FOOD & BEVERAGES" -> "FOOD & BEVERAGE".
KNOWN_SECTORS = [
    "TEXTILE & GARMENT", "AUTOMOTIVE", "FOOD & BEVERAGE", "OIL & GAS", "PROPERTY", "CONSTRUCTION",
    "TELECOMMUNICATIONS", "LOGISTIC & TRANSPORTATION", "ENGINEERING", "COMMODITY", "RESOURCES", "HEALTH",
    "GENERAL COMMERCE & MANUFACTURE", "METAL", "CHEMICALS & PLASTICS", "SERVICES", "FURNITURE", "CONSUMER",
    "ELECTRICAL & ELECTRONICS", "FINANCIAL INSTITUTION", "FISHERY AND FARMING", "PAPER", "PUBLIC UTILITIES",
]
SECTOR_MATCH_CUTOFF = 0.86   # how close a misspelt sector name must be to count as the same sector
POSITIONS = ["Upstream", "Midstream", "Downstream", "Enabling"]

# Column names: the first name found in the file is used (case, spaces, "-" and "_" are ignored).
COLUMN_ALIASES = {
    "customerid": ["customerid", "cif", "cifno", "customer_id"],
    "customername": ["customername", "name"],
    "customercompany": ["customercompany", "company"],
    "subsegment": ["subsegment"],
    "region": ["region"],
    "area": ["area"],
    "officename": ["officename", "branch"],
    "has_pt_ocbc": ["has_pt_ocbc", "has_pt", "haspt"],
    "main_sector": ["Sector", "main_sector", "economicsector_segment"],   # "1600000000-SERVICES" -> "SERVICES"
    "product": ["economicsector_subsubsegment2", "economicsector_subsubsegment1", "economicsector_subsegment1", "economicsector_subsegment"],
    "value_chain_supplier": ["value_chain_supplier", "value_chain_supp", "value_chain_sup"],
    "value_chain_buyer": ["value_chain_buyer"],
    "econ_bicode": ["economicsector_bicode", "economicsector_biocode"],
    "econ_subsegment": ["economicsector_subsegment"],
    "upstream": ["Upstream (sectors)", "upstream", "upstream_sectors"],
    "midstream": ["Mid-stream (sectors)", "midstream", "mid_stream", "midstream_sectors"],
    "downstream": ["Downstream (sectors)", "downstream", "downstream_sectors"],
    "enabling": ["Enabling", "enabling", "enabling (sectors)", "enabling_sectors"],
    "sto": ["sto"],
    "sto_annum": ["sto_annum"],
    "aum_indiv": ["aum_indiv", "aum"],
    "lum_indiv": ["lum_indiv", "lum"],
    "bizvolume": ["bizvolume", "biz_volume", "business_volume"],
    "avrg_trx_3m": ["avrg_trx_3m"],
    "total_PT": ["total_PT", "total_pt"],
}
REQUIRED = ["customerid", "main_sector", "has_pt_ocbc", "upstream", "midstream", "downstream", "enabling"]
NUMERIC = ["sto", "sto_annum", "aum_indiv", "lum_indiv", "bizvolume", "avrg_trx_3m", "total_PT"]
# how customers are ranked (biggest first); biz volume = AUM + LUM
RANK_OPTIONS = {"Biz volume (AUM + LUM)": "bizvolume", "AUM": "aum_indiv", "LUM": "lum_indiv"}
# how partner candidates are ranked (biggest first); ties go to the second measure
PARTNER_RANK_OPTIONS = {"Annual turnover (sto_annum)": "sto_annum", "Biz volume (AUM + LUM)": "bizvolume",
                        "AUM": "aum_indiv", "LUM": "lum_indiv"}


def tie_break(prank: str) -> str:
    return "bizvolume" if prank == "sto_annum" else "sto_annum"

# chain columns and how they are shown
CHAIN_COLS = {"up": "upstream", "mid": "midstream", "down": "downstream", "en": "enabling"}
REL_LABEL = {"up": "Supplier (upstream)", "down": "Buyer (downstream)", "mid": "Two-way (midstream)", "en": "Enabling"}
POS_LABEL = {"up": "Upstream", "mid": "Midstream", "down": "Downstream", "en": "Enabling"}

# matchmaking: candidates per relation in Match results / Customer detail
TOP_PER_RELATION = 5
TOP_IN_DETAIL = 20
LIST_POSITION_PENALTY = 0.06  # each later-listed sector scores 6% lower (max 50%)

# cell values that mean "empty" (Excel formulas often leave 0 or #N/A)
NULL_TEXT = {"", "-", "--", "0", "0.0", "null", "nan", "n/a", "na", "n.a", "none", "undefined",
             "#n/a", "#value!", "#ref!", "#div/0!", "#name?", "#null!", "#num!"}

COLOR_PT = {"Has a PT": "#2a6fdb", "No PT": "#c4475c"}
COLOR_POS = {"up": "#d9822b", "mid": "#2e9e6a", "down": "#2a6fdb", "en": "#7a5bd1"}
PALETTE = ["#2a6fdb", "#d9822b", "#2e9e6a", "#c4475c", "#7a5bd1", "#3aa3b8", "#b58b00", "#8a8a8a"]


# ---------------------------------------------------------------------------
# Loading and cleaning
# ---------------------------------------------------------------------------

def clean_text(v) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def id_text(v) -> str:
    if isinstance(v, float) and not math.isnan(v) and v.is_integer():
        return str(int(v))
    return clean_text(v)


def is_null(v) -> bool:
    return clean_text(v).lower() in NULL_TEXT


def strip_code(v) -> str:
    """'1600000000-SERVICES' -> 'SERVICES'; '8164 - Perdagangan ...' -> 'Perdagangan ...'."""
    return re.sub(r"^\s*\d{2,}\s*-\s*", "", clean_text(v))


def col_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def resolve_columns(columns: list[str]) -> dict[str, str]:
    """Canonical name -> the file's column name."""
    by_key = {}
    for c in columns:
        by_key.setdefault(col_key(c), c)
    out = {}
    for canon, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if col_key(a) in by_key:
                out[canon] = by_key[col_key(a)]
                break
    return out


def sector_key(name: str) -> str:
    t = clean_text(name).upper().replace(" AND ", " & ")
    words = re.findall(r"[A-Z0-9&]+", t)
    return " ".join(w[:-1] if len(w) > 3 and w.endswith("S") else w for w in words)


class SectorNames:
    """Maps any spelling of a sector to one display name."""

    def __init__(self, main_values: list[str]):
        counts: dict[str, dict[str, int]] = {}
        for v in main_values:
            if v:
                counts.setdefault(sector_key(v), {}).setdefault(v, 0)
                counts[sector_key(v)][v] += 1
        # the spelling used most in the main-sector column wins; known names fill the gaps
        self.names = {k: max(d, key=d.get) for k, d in counts.items()}
        for s in KNOWN_SECTORS:
            self.names.setdefault(sector_key(s), s)
        self.cache: dict[str, str] = {}

    def __call__(self, v) -> str:
        t = clean_text(v)
        if t in self.cache:
            return self.cache[t]
        k = sector_key(t)
        if k in self.names:
            out = self.names[k]
        else:
            near = difflib.get_close_matches(k, list(self.names), n=1, cutoff=SECTOR_MATCH_CUTOFF)
            out = self.names[near[0]] if near else t.upper()
        self.cache[t] = out
        return out


def parse_list(v, names: SectorNames) -> list[str]:
    """One cell of upstream/midstream/downstream/enabling -> list of sectors (split on ; , | or new line)."""
    out: list[str] = []
    if v is None or (isinstance(v, (int, float)) and not isinstance(v, bool) and (v == 0 or v != v)):
        return out
    for part in re.split(r"[;,|\n]", clean_text(v)):
        t = clean_text(part)
        if not t or is_null(t):
            continue
        c = names(t)
        if c not in out:
            out.append(c)
    return out


def to_number(v) -> float:
    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
        return float(v)
    t = clean_text(v).replace(",", "")
    if not t or is_null(t) and t not in ("0", "0.0"):
        return float("nan")
    try:
        return float(t)
    except ValueError:
        return float("nan")


@st.cache_data(show_spinner="Reading the file…")
def load_file(content: bytes, name: str) -> tuple[pd.DataFrame, list[str], str]:
    """Read the uploaded Excel/CSV. Returns (raw data, original column order, sheet used).
    Excel: the sheet named 'data' is used, otherwise the first sheet."""
    if name.lower().endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(content), dtype=str, keep_default_na=False, encoding="utf-8-sig")
        sheet = "CSV"
    else:
        xls = pd.ExcelFile(io.BytesIO(content))
        sheet = next((s for s in xls.sheet_names if s.strip().lower() == "data"), xls.sheet_names[0])
        raw = pd.read_excel(xls, sheet_name=sheet, dtype=object)
    raw.columns = [clean_text(c) for c in raw.columns]
    raw = raw.loc[:, [c for c in raw.columns if c and not c.lower().startswith("unnamed")]]
    cols = resolve_columns(list(raw.columns))
    if "customerid" in cols:
        raw = raw[raw[cols["customerid"]].map(id_text) != ""]
    return raw.reset_index(drop=True), list(raw.columns), sheet


@st.cache_data(show_spinner=False)
def prepare(data_key: str, _raw: pd.DataFrame) -> pd.DataFrame:
    """Build the working table under canonical column names and parse the chain columns once."""
    raw = _raw
    cols = resolve_columns(list(raw.columns))
    get = lambda canon: raw[cols[canon]] if canon in cols else pd.Series([""] * len(raw), index=raw.index)
    df = pd.DataFrame(index=raw.index)
    df["customerid"] = get("customerid").map(id_text)
    for c in ["customername", "customercompany", "subsegment", "region", "area", "officename"]:
        df[c] = get(c).map(lambda v: "" if is_null(v) and clean_text(v) not in ("0",) else clean_text(v))
    df["customername"] = df["customername"].where(df["customername"] != "", df["customerid"])
    df["product"] = get("product").map(strip_code)
    for c in ["value_chain_supplier", "value_chain_buyer"]:
        df[c] = get(c).map(lambda v: "(blank)" if is_null(v) else clean_text(v).capitalize())
    for c in NUMERIC:
        df[c] = get(c).map(to_number)
    # biz volume = AUM + LUM (a bizvolume column in the file is used where it has a value)
    calc = (df["aum_indiv"].fillna(0) + df["lum_indiv"].fillna(0)).where(df["aum_indiv"].notna() | df["lum_indiv"].notna())
    df["bizvolume"] = df["bizvolume"].where(df["bizvolume"].notna(), calc)
    pt = get("has_pt_ocbc").map(lambda v: clean_text(v).upper())
    df["has_pt_ocbc"] = pt.isin(["Y", "YES", "1", "TRUE"]).map({True: "Y", False: "N"})
    main = get("main_sector").map(strip_code)
    names = SectorNames(main.tolist())
    df["economicsector_segment"] = main.map(lambda v: names(v) if v and not is_null(v) else "(no sector)")
    for key, col in CHAIN_COLS.items():
        df["list_" + key] = get(col).map(lambda v: parse_list(v, names))
    for c in ["econ_bicode", "econ_subsegment"]:
        df[c] = get(c).map(lambda v: "" if is_null(v) else clean_text(v))
    # mapping check: which parts are empty for this customer
    df["no_chain"] = ~df[["list_" + k for k in CHAIN_COLS]].apply(lambda r: any(len(x) for x in r), axis=1)
    df["no_vc_supplier"] = df["value_chain_supplier"] == "(blank)"
    df["no_vc_buyer"] = df["value_chain_buyer"] == "(blank)"
    df["no_sector"] = df["economicsector_segment"] == "(no sector)"
    df["label"] = (df["customerid"] + " · " + df["customername"] + " · "
                   + df["customercompany"].where((df["has_pt_ocbc"] == "Y") & (df["customercompany"] != ""),
                                                 df["has_pt_ocbc"].map({"Y": "has PT", "N": "no PT"})))
    return df


# ---------------------------------------------------------------------------
# Matchmaking logic (same rules as the web dashboard)
# ---------------------------------------------------------------------------

def chain_of(sector: str, df: pd.DataFrame) -> tuple[int, pd.DataFrame]:
    """Combined chain of a sector over every customer in it.
    Each linked sector goes to the column it is listed in most often."""
    in_sec = df[df["economicsector_segment"] == sector]
    n = len(in_sec)
    counts: dict[str, dict[str, int]] = {}
    for key in CHAIN_COLS:
        for lst in in_sec["list_" + key]:
            for name in lst:
                counts.setdefault(name, {"up": 0, "mid": 0, "down": 0, "en": 0})[key] += 1
    rows = []
    for name, c in counts.items():
        pos = max(["up", "mid", "down", "en"], key=lambda k: (c[k], -["up", "mid", "down", "en"].index(k)))
        rows.append({"linked_sector": name, "pos": pos, "customers_listing": c[pos],
                     "share": c[pos] / n if n else 0.0})
    links = pd.DataFrame(rows, columns=["linked_sector", "pos", "customers_listing", "share"])
    links = links.sort_values(["customers_listing", "linked_sector"], ascending=[False, True]).reset_index(drop=True)
    return n, links


def customer_chain(c, df: pd.DataFrame) -> tuple[dict[str, list[str]], bool]:
    """The customer's own lists; the sector's combined chain when all four are empty.
    `c` is one customer (a row or a dict)."""
    own = {k: list(c["list_" + k]) for k in CHAIN_COLS}
    if any(own.values()):
        return own, False
    _, links = chain_of(c["economicsector_segment"], df)
    return {k: links.loc[links["pos"] == k, "linked_sector"].tolist() for k in CHAIN_COLS}, True


@st.cache_data(show_spinner=False)
def pt_pool(data_key: str, _df: pd.DataFrame, prank: str = "sto_annum") -> dict:
    """PT owners as plain dicts, biggest first by `prank` (ties by the second measure):
    by sector, and by (sector, region)."""
    pool = _df[_df["has_pt_ocbc"] == "Y"].sort_values([prank, tie_break(prank)], ascending=False, na_position="last")
    by_sector: dict[str, list[dict]] = {}
    by_region: dict[tuple[str, str], list[dict]] = {}
    for r in pool.to_dict("records"):
        by_sector.setdefault(r["economicsector_segment"], []).append(r)
        by_region.setdefault((r["economicsector_segment"], r["region"]), []).append(r)
    return {"sector": by_sector, "region": by_region}


def candidates(c, key: str, k: int, df: pd.DataFrame, pool: dict, lists: dict | None = None,
               prank: str = "sto_annum") -> list[dict]:
    """PT owners in the sectors the customer lists under `key`, biggest first by `prank`
    (annual turnover, biz volume, AUM or LUM), earlier-listed sectors slightly preferred,
    ties broken by the second measure."""
    second = tie_break(prank)
    num = lambda v: v if v == v and v is not None else 0  # NaN-safe
    if lists is None:
        lists, _ = customer_chain(c, df)
    found: dict[str, dict] = {}
    for i, sec in enumerate(lists[key]):
        pick = pool["sector"].get(sec, [])[: k + 1]
        for r in pick:
            if r["customerid"] == c["customerid"] or r["customerid"] in found:
                continue
            same = r["region"] == c["region"]
            score = math.log10(max(1.0, num(r[prank]))) * (1 - min(0.5, LIST_POSITION_PENALTY * i))
            found[r["customerid"]] = {"r": r, "listed_position": i + 1, "same_region": same, "score": score,
                                      "tie": num(r[second])}
    return sorted(found.values(), key=lambda x: (-x["score"], -x["tie"]))[:k]


@st.cache_data(show_spinner="Finding partner matches…")
def match_results(data_key: str, focus_ids: tuple[str, ...], _df: pd.DataFrame, rank_col: str = "bizvolume",
                  prank: str = "sto_annum") -> pd.DataFrame:
    df = _df
    pool = pt_pool(data_key, df, prank)
    focus = df[df["customerid"].isin(set(focus_ids))].sort_values(rank_col, ascending=False, na_position="last")
    focus = focus.assign(customer_rank=focus[rank_col].rank(ascending=False, method="min", na_option="bottom"))
    fallback_cache: dict[str, dict] = {}
    out = []
    for c in focus.to_dict("records"):
        own = {k: c["list_" + k] for k in CHAIN_COLS}
        if any(own.values()):
            lists = own
        else:
            sec = c["economicsector_segment"]
            if sec not in fallback_cache:
                fallback_cache[sec] = customer_chain(c, df)[0]
            lists = fallback_cache[sec]
        for key in ["up", "down", "mid", "en"]:
            for rank, x in enumerate(candidates(c, key, TOP_PER_RELATION, df, pool, lists, prank), 1):
                r = x["r"]
                out.append({
                    "customer_rank": int(c["customer_rank"]), rank_col: c[rank_col],
                    "customerid": c["customerid"], "customername": c["customername"],
                    "customercompany": c["customercompany"], "subsegment": c["subsegment"],
                    "economicsector_segment": c["economicsector_segment"], "region": c["region"],
                    "officename": c["officename"], "relation": REL_LABEL[key], "partner_rank": rank,
                    "sector_listed_position": x["listed_position"],
                    "partner_customerid": r["customerid"], "partner_customername": r["customername"],
                    "partner_company": r["customercompany"], "partner_sector": r["economicsector_segment"],
                    "partner_product": r.get("product", ""), "partner_region": r["region"],
                    "partner_officename": r["officename"], "same_region": "Y" if x["same_region"] else "N",
                    "partner_sto_annum": r["sto_annum"], "partner_bizvolume": r["bizvolume"],
                    "partner_aum_indiv": r["aum_indiv"], "partner_lum_indiv": r["lum_indiv"],
                    "partner_ranked_by": prank, "score": round(x["score"], 3),
                })
    return pd.DataFrame(out)


@st.cache_data(show_spinner=False)
def sector_pairs(data_key: str, focus_ids: tuple[str, ...], _df: pd.DataFrame) -> pd.DataFrame:
    """Supplier -> buyer sector pairs from the upstream/downstream columns."""
    sub = _df[_df["customerid"].isin(set(focus_ids))]
    mentions: dict[tuple[str, str], set] = {}
    for r in sub.to_dict("records"):
        x = r["economicsector_segment"]
        for y in r["list_down"]:
            if y != x:
                mentions.setdefault((x, y), set()).add(r["customerid"])
        for y in r["list_up"]:
            if y != x:
                mentions.setdefault((y, x), set()).add(r["customerid"])
    # links come from the filtered customers; PT owners on each side are counted over the whole
    # portfolio (partners can be anyone), so a filter on one sector still shows its pairs
    pt = _df[_df["has_pt_ocbc"] == "Y"].groupby("economicsector_segment").size()
    rows = []
    for (s, b), ids in mentions.items():
        n_s, n_b = int(pt.get(s, 0)), int(pt.get(b, 0))
        if n_s and n_b:
            rows.append({"supplier_sector": s, "buyer_sector": b, "mentions": len(ids),
                         "supplier_pt_owners": n_s, "buyer_pt_owners": n_b, "combinations": n_s * n_b})
    cols = ["supplier_sector", "buyer_sector", "mentions", "supplier_pt_owners", "buyer_pt_owners", "combinations"]
    return pd.DataFrame(rows, columns=cols).sort_values(["combinations", "mentions"], ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Small helpers for display
# ---------------------------------------------------------------------------

def rp(v) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    a = abs(v)
    if a >= 1e12:
        return f"Rp {v / 1e12:,.1f} T"
    if a >= 1e9:
        return f"Rp {v / 1e9:,.1f} B"
    if a >= 1e6:
        return f"Rp {v / 1e6:,.0f} M"
    return f"Rp {v:,.0f}"


def csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")  # BOM so Excel opens it with the right encoding


def download(label: str, df: pd.DataFrame, name: str, key: str):
    st.download_button(label, csv_bytes(df), file_name=f"{name}_{date.today():%Y-%m-%d}.csv",
                       mime="text/csv", key=key, disabled=df.empty)


def chips(title: str, values: list[str], limit: int = 15) -> str:
    shown = values[:limit]
    more = f" … +{len(values) - limit:,} more" if len(values) > limit else ""
    return f"**{title}:** " + "  ".join(f"`{v}`" for v in shown) + more


def hbar(data: pd.DataFrame, cat: str, color: str, colors: dict, order: list[str] | None = None, legend_columns: int | None = None):
    """Horizontal stacked bars with a total at the end of each bar, roomy spacing and the legend below."""
    if data.empty:
        return None
    totals = data.groupby(cat)["count"].sum()
    if order:
        cats = [c for c in order if c in totals.index] + [c for c in totals.sort_values(ascending=False).index if c not in order]
    else:
        cats = totals.sort_values(ascending=False).index.tolist()
    domain = [d for d in colors if d in set(data[color])] or list(colors)
    y = alt.Y(f"{cat}:N", sort=cats, title=None,
              scale=alt.Scale(paddingInner=0.38, paddingOuter=0.25),
              axis=alt.Axis(labelLimit=190, labelPadding=10, ticks=False, domain=False, labelFontSize=12, labelColor="#3d3d3a"))
    xmax = float(totals.max()) * 1.15 or 1
    x = alt.X("count:Q", title=None, stack="zero", scale=alt.Scale(domain=[0, xmax], nice=False),
              axis=alt.Axis(tickCount=4, format=",d", tickMinStep=1, grid=True, gridColor="#ecebe6", domain=False, ticks=False,
                            labelColor="#8a8983", labelFontSize=11))
    bars = alt.Chart(data).mark_bar(cornerRadiusEnd=3).encode(
        y=y, x=x,
        color=alt.Color(f"{color}:N", scale=alt.Scale(domain=domain, range=[colors[d] for d in domain]),
                        legend=None if legend_columns == 0 else alt.Legend(orient="bottom", direction="horizontal", title=None, labelLimit=220,
                                          symbolType="square", symbolSize=110, labelFontSize=12, columnPadding=16,
                                          )),
        order=alt.Order(f"{color}_order:Q"),
        tooltip=[alt.Tooltip(f"{cat}:N", title="Group"), alt.Tooltip(f"{color}:N", title="Split"),
                 alt.Tooltip("count:Q", title="Customers", format=",")],
    )
    tot = pd.DataFrame({cat: totals.index, "total": totals.values})
    labels = alt.Chart(tot).mark_text(align="left", dx=6, fontSize=12, color="#6b6a66").encode(
        y=alt.Y(f"{cat}:N", sort=cats), x=alt.X("total:Q"), text=alt.Text("total:Q", format=","))
    return (alt.layer(bars, labels)
            .properties(height=alt.Step(38), padding={"left": 4, "right": 12, "top": 8, "bottom": 4})
            .configure_view(stroke=None)
            .configure_legend(padding=10, offset=14))


def chart_card(title: str, note: str, chart, legend: dict | None = None):
    """A bordered card; `legend` (label -> colour) is written under the chart when the chart has none of its own."""
    with st.container(border=True):
        st.markdown(f"**{title}**")
        st.caption(note)
        if chart is None:
            st.caption("No customers under these filters.")
        else:
            st.altair_chart(chart, use_container_width=True)
            if legend:
                st.markdown(" &nbsp; ".join(
                    f"<span style='white-space:nowrap;font-size:0.8rem;color:#6b6a66'>"
                    f"<span style='display:inline-block;width:10px;height:10px;background:{c};border-radius:2px;margin-right:5px'></span>{html.escape(k)}</span>"
                    for k, c in legend.items()), unsafe_allow_html=True)


def counts_by(df: pd.DataFrame, cat: str, split: str, split_order: list[str]) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[cat, split, "count", f"{split}_order"])
    g = df.groupby([cat, split]).size().reset_index(name="count")
    g[f"{split}_order"] = g[split].map({s: i for i, s in enumerate(split_order)})
    return g


def chain_dot(sector: str, n: int, links: pd.DataFrame, pt_counts: pd.Series) -> str:
    """Graphviz DOT for the chain map: upstream -> sector -> downstream, midstream two-way."""
    def q(s):
        return '"' + s.replace('"', "'") + '"'

    def label(name, sub):
        return q(f"{name}\n{sub}")

    prod = links[(links["pos"] != "en") & (links["linked_sector"] != sector)]
    max_share = prod["share"].max() if not prod.empty else 1
    width = lambda s: f"{1.2 + 3.3 * math.sqrt(s / max_share if max_share else 0):.2f}"
    lines = [
        "digraph G {", "rankdir=LR; splines=true; nodesep=0.25; ranksep=1.1; bgcolor=transparent;",
        'node [shape=box, style="rounded,filled", fillcolor="#ffffff", color="#d9d6cf", fontname="Helvetica", fontsize=11, margin="0.18,0.08"];',
        'edge [arrowsize=0.7];',
        f'{q("__sel")} [label={label(sector, f"{n:,} customers · {int(pt_counts.get(sector, 0)):,} with a PT")}, '
        'fillcolor="#f2f1ec", color="#1f1f1e", penwidth=2, fontsize=12];',
    ]
    groups = {k: prod[prod["pos"] == k] for k in ["up", "mid", "down"]}
    for key, g in groups.items():
        for _, l in g.iterrows():
            nid = q(f"{key}:{l['linked_sector']}")
            sub = f"{l['share']:.0%} listing · {int(pt_counts.get(l['linked_sector'], 0)):,} PT"
            lines.append(f'{nid} [label={label(l["linked_sector"], sub)}, color="{COLOR_POS[key]}"];')
            if key == "up":
                lines.append(f'{nid} -> {q("__sel")} [color="{COLOR_POS[key]}", penwidth={width(l["share"])}];')
            elif key == "down":
                lines.append(f'{q("__sel")} -> {nid} [color="{COLOR_POS[key]}", penwidth={width(l["share"])}];')
            else:
                lines.append(f'{q("__sel")} -> {nid} [dir=both, color="{COLOR_POS[key]}", penwidth={width(l["share"])}];')
    if not groups["mid"].empty:
        # midstream sits in the same column as the selected sector, under it, drawn dashed
        mids = [q(f"mid:{s}") for s in groups["mid"]["linked_sector"]]
        lines.append("{ rank=same; " + q("__sel") + "; " + "; ".join(mids) + "; }")
        lines.append(" ".join(f'{m} [style="rounded,dashed,filled"];' for m in mids))
    lines.append("}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Business Matchmaking Dashboard", layout="wide")
st.title("Business Matchmaking Dashboard")

with st.sidebar:
    st.header("Data")
    up = st.file_uploader("Customer data (Excel or CSV)", type=["xlsx", "xls", "csv"],
                          help="Use the template: one row per customer, sheet named 'data'.")

if up is None:
    st.info("Upload the customer data file in the sidebar to start. "
            "Use **business_matchmaking_template.xlsx** for the expected columns.")
    st.stop()

raw, raw_columns, sheet_used = load_file(up.getvalue(), up.name)
found_cols = resolve_columns(raw_columns)
missing = [c for c in REQUIRED if c not in found_cols]
if missing:
    st.error("These columns are missing from the file: " + ", ".join(missing)
             + ". Accepted names: " + "; ".join(f"{c} = {' / '.join(COLUMN_ALIASES[c])}" for c in missing))
    st.stop()
import hashlib
data_key = hashlib.md5(up.getvalue()).hexdigest()
df = prepare(data_key, raw)
pt_counts_all = df[df["has_pt_ocbc"] == "Y"].groupby("economicsector_segment").size()

def ordered(values, order=None):
    vals = sorted({v for v in values if v})
    if order:
        return [v for v in order if v in vals] + [v for v in vals if v not in order]
    return vals

with st.sidebar:
    st.caption(f"Loaded {len(df):,} customers from {up.name} ({sheet_used}).")

# ---------------------------------------------------------------- Filters
# Picked values are kept in session state under "keep_<name>", apart from the widgets, so they stay
# picked when you switch filter mode or narrow another filter, and saved selections can restore them.
FILTERS = {  # name: (label, column)
    "region": ("Region", "region"),
    "area": ("Area", "area"),
    "office": ("Branch", "officename"),
    "sub": ("Subsegment", "subsegment"),
    "kat": ("Main sector", "economicsector_segment"),
    "vc": ("Value chain (supplier)", "value_chain_supplier"),
    "vcb": ("Value chain (buyer)", "value_chain_buyer"),
    "cust": ("Customer ID", "customerid"),
}
MODE_FILTERS = {
    "Sector": ["region", "area", "sub", "kat", "vc", "vcb"],
    "Customer": ["region", "area", "office", "sub", "cust"],
}
ss = st.session_state
ss.setdefault("mode", "Sector")
ss.setdefault("presets", {})
for name in FILTERS:
    ss.setdefault("keep_" + name, [])


def picked(name: str) -> list:
    return ss["keep_" + name]


def _sync(name: str, options: list):
    new = ss["w_" + name]
    hidden = [v for v in ss["keep_" + name] if v not in options]  # picked earlier, not offered right now
    ss["keep_" + name] = list(dict.fromkeys(new + hidden))


def multi(name: str, options: list, *, format_func=str, placeholder=None, container=st):
    label = FILTERS[name][0]
    ss["w_" + name] = [v for v in picked(name) if v in set(options)]
    return container.multiselect(label, options, key="w_" + name, format_func=format_func,
                                 on_change=_sync, args=(name, options),
                                 placeholder=placeholder or f"All {label.lower()}s")


def narrowed(exclude: str) -> pd.DataFrame:
    """Customers matching every active filter of this mode except one (used to narrow that filter's options)."""
    m = pd.Series(True, index=df.index)
    for name in MODE_FILTERS[ss["mode"]]:
        if name != exclude and picked(name):
            m &= df[FILTERS[name][1]].isin(picked(name))
    return df[m]


def reset_filters():
    for name in FILTERS:
        ss["keep_" + name] = []


def save_preset():
    nm = clean_text(ss.get("preset_name", ""))
    if nm:
        ss["presets"][nm] = {"mode": ss["mode"], **{n: list(picked(n)) for n in FILTERS}}
        ss["preset_name"] = ""


def apply_preset(nm: str):
    p = ss["presets"].get(nm)
    if p:
        ss["mode"] = p.get("mode", "Sector")
        for n in FILTERS:
            ss["keep_" + n] = list(p.get(n, []))


id_label = dict(zip(df["customerid"], df["label"]))
ss.setdefault("rank_by", next(iter(RANK_OPTIONS)))
if ss["rank_by"] not in RANK_OPTIONS:
    ss["rank_by"] = next(iter(RANK_OPTIONS))
rank_col = RANK_OPTIONS[ss["rank_by"]]
rank_name = ss["rank_by"].split(" (")[0]
ss.setdefault("partner_rank_by", next(iter(PARTNER_RANK_OPTIONS)))
if ss["partner_rank_by"] not in PARTNER_RANK_OPTIONS:
    ss["partner_rank_by"] = next(iter(PARTNER_RANK_OPTIONS))
prank = PARTNER_RANK_OPTIONS[ss["partner_rank_by"]]
prank_name = {"sto_annum": "annual turnover", "bizvolume": "biz volume", "aum_indiv": "AUM", "lum_indiv": "LUM"}[prank]
id_order = df.sort_values(rank_col, ascending=False, na_position="last")["customerid"].tolist()

n_unmapped = int((df["no_chain"] | df["no_vc_supplier"] | df["no_vc_buyer"] | df["no_sector"]).sum())
# tab order on screen: Sector chain explorer, Matchmaking, Unmapped data
tab_s, tab_c, tab_u = st.tabs(["Sector chain explorer", "Matchmaking", f"Unmapped data ({n_unmapped:,})"])

with tab_c, st.container(border=True):
    top = st.columns([2.6, 2.4, 0.9, 2.4, 0.9], vertical_alignment="bottom")
    with top[0]:
        def _set_mode():
            if ss.get("w_mode"):
                ss["mode"] = ss["w_mode"]
        ss["w_mode"] = ss["mode"]
        seg = getattr(st, "segmented_control", None)  # Streamlit 1.40+; older versions get radio buttons
        if seg:
            seg("Filter by", ["Sector", "Customer"], key="w_mode", on_change=_set_mode)
        else:
            st.radio("Filter by", ["Sector", "Customer"], horizontal=True, key="w_mode", on_change=_set_mode)
    mode = ss["mode"]
    top[1].text_input("Save this selection as", key="preset_name", placeholder="e.g. Jabodetabek F&B")
    top[2].button("Save", on_click=save_preset, use_container_width=True, help="Save the current filters under this name")
    names = list(ss["presets"])
    sel = top[3].selectbox("Saved selections", names, index=None,
                           placeholder="None saved yet" if not names else f"{len(names)} saved · pick one", key="preset_pick")
    top[4].button("Reset", on_click=reset_filters, use_container_width=True, help="Clear every filter")
    if sel:
        a1, a2, _ = st.columns([1, 1, 6])
        a1.button("Apply", on_click=apply_preset, args=(sel,), use_container_width=True)
        if a2.button("Delete", use_container_width=True):
            ss["presets"].pop(sel, None)
            st.rerun()

    if mode == "Sector":
        r1 = st.columns(3)
        multi("region", ordered(narrowed("region")["region"]), container=r1[0])
        multi("area", ordered(narrowed("area")["area"]), container=r1[1])
        multi("sub", ordered(narrowed("sub")["subsegment"]), container=r1[2])
        r2 = st.columns(3)
        multi("kat", ordered(narrowed("kat")["economicsector_segment"]), container=r2[0], placeholder="All sectors")
        multi("vc", ordered(narrowed("vc")["value_chain_supplier"], POSITIONS + ["(blank)"]), container=r2[1], placeholder="All positions")
        multi("vcb", ordered(narrowed("vcb")["value_chain_buyer"], POSITIONS + ["(blank)"]), container=r2[2], placeholder="All positions")
        st.caption("Pick one or more values in each filter; within a filter any picked value matches, across filters all must match. "
                   "Each list only offers values that fit the other filters. Leave a filter empty to keep everyone.")
    else:
        r1 = st.columns(4)
        multi("region", ordered(narrowed("region")["region"]), container=r1[0])
        multi("area", ordered(narrowed("area")["area"]), container=r1[1])
        multi("office", ordered(narrowed("office")["officename"]), container=r1[2], placeholder="All branches")
        multi("sub", ordered(narrowed("sub")["subsegment"]), container=r1[3])
        pool_ids = set(narrowed("cust")["customerid"])
        multi("cust", [i for i in id_order if i in pool_ids], format_func=lambda i: id_label.get(i, i),
              placeholder="All customers — type a CIF or name to search")
        with st.expander("Add many customers at once"):
            e1, e2 = st.columns(2)
            pasted = e1.text_area("Paste customer IDs (one per line, or separated by comma / space)", key="paste_ids", height=110)
            if e1.button("Add pasted IDs"):
                want = [t for t in re.split(r"[\s,;]+", pasted) if t]
                known = [t for t in want if t in id_label]
                ss["keep_cust"] = list(dict.fromkeys(picked("cust") + known))
                unknown = [t for t in want if t not in id_label]
                ss["paste_msg"] = (f"Added {len(known):,} customer(s)."
                                   + (f" {len(unknown):,} ID(s) not in the file: " + ", ".join(unknown[:15]) + (" …" if len(unknown) > 15 else "") if unknown else ""))
                st.rerun()
            if ss.get("paste_msg"):
                (e1.warning if "not in the file" in ss["paste_msg"] else e1.success)(ss.pop("paste_msg"))
            q = e2.text_input("…or every customer whose CIF, name, sector, region or branch contains", key="cust_query",
                              placeholder="e.g. jabodetabek, food, 10001")
            if e2.button("Add all matches", disabled=not q.strip()):
                hay = (df["label"] + " " + df["economicsector_segment"] + " " + df["region"] + " " + df["area"] + " " + df["officename"]).str.lower()
                base = narrowed("cust")
                add = [i for i in df.loc[hay.str.contains(re.escape(q.strip().lower())), "customerid"] if i in set(base["customerid"])]
                ss["keep_cust"] = list(dict.fromkeys(picked("cust") + add))
                st.rerun()
        st.caption("Region, area, branch and subsegment narrow the customer list; picked customers are then the only ones shown. "
                   "With no customer picked, everyone matching the other filters is shown.")
    rk = st.columns([1.4, 1.4, 2.2], vertical_alignment="bottom")
    rk[0].selectbox("Rank customers by", list(RANK_OPTIONS), key="rank_by",
                    help="Biggest first. Sets the order of customers in every table and list, and the rank number.")
    rk[1].selectbox("Rank partners by", list(PARTNER_RANK_OPTIONS), key="partner_rank_by",
                    help="Biggest first. Decides which partner candidates make the top list and their partner_rank. "
                         "Ties go to biz volume (or to annual turnover when ranking by AUM / LUM).")
    rk[2].caption("Rank 1 = biggest. Customers are ranked within the current filter; partners come from the "
                  "whole portfolio. Biz volume = AUM + LUM.")

# apply filters of the current mode
mask = pd.Series(True, index=df.index)
for name in MODE_FILTERS[mode]:
    if picked(name):
        mask &= df[FILTERS[name][1]].isin(picked(name))
fdf = df[mask]
fdf = fdf.assign(customer_rank=fdf[rank_col].rank(ascending=False, method="min", na_option="bottom").astype(int))
picked_summary = []
for name in MODE_FILTERS[mode]:
    vals = picked(name)
    if vals:
        shown = [id_label.get(v, v).split(" · ")[0] + " · " + id_label.get(v, v).split(" · ")[1] if name == "cust" and v in id_label else v for v in vals]
        picked_summary.append((FILTERS[name][0], shown))

focus_ids = tuple(fdf["customerid"])


# ---------------------------------------------------------------- Customer tab
with tab_c:
    # what is selected
    box = st.container(border=True)
    box.markdown(f"**Showing {len(fdf):,} of {len(df):,} customers** · filter by {mode.lower()}"
                 + ("" if picked_summary else " · no filters picked, so everyone is shown"))
    for title, vals in picked_summary:
        box.markdown(chips(title, vals, limit=25))
    if mode == "Customer" and picked("cust"):
        outside = len(set(picked("cust")) - set(fdf["customerid"]))
        if outside:
            box.caption(f"{outside:,} picked customer(s) are outside the region / area / branch / subsegment filters, "
                        "so they are kept but not shown. Clear those filters to see them.")
    n_nochain_f = int(fdf["no_chain"].sum())
    if n_nochain_f:
        box.caption(f"{n_nochain_f:,} of these customers have empty Upstream / Mid-stream / Downstream / Enabling columns; "
                    "the combined chain of their sector is used for them. See the Unmapped data tab for the list.")
    if picked_summary and not len(fdf):
        box.warning("No customer matches every filter. Remove a value above or press Reset.")

    n_pt = int((fdf["has_pt_ocbc"] == "Y").sum())
    matches = match_results(data_key, focus_ids, df, rank_col, prank)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Customers", f"{len(fdf):,}")
    k2.metric("Customers with a PT" + (f" ({n_pt / len(fdf):.1%})" if len(fdf) else ""), f"{n_pt:,}")
    k3.metric("Partner matches found", f"{len(matches):,}")
    k4.metric("Customers ranked by", rank_name)
    k5, k6, k7, k8 = st.columns(4)
    k5.metric("Total biz volume (AUM + LUM)", rp(fdf["bizvolume"].sum()))
    k6.metric("Total AUM", rp(fdf["aum_indiv"].sum()))
    k7.metric("Total LUM", rp(fdf["lum_indiv"].sum()))
    k8.metric("Annual company turnover", rp(fdf["sto_annum"].sum()))

    pt_split = fdf.assign(pt=fdf["has_pt_ocbc"].map({"Y": "Has a PT", "N": "No PT"}))
    st.write("")
    c1, c2, c3 = st.columns(3, gap="medium")
    with c1:
        chart_card("Value chain as supplier", "Customers by their position when selling.",
                   hbar(counts_by(pt_split, "value_chain_supplier", "pt", list(COLOR_PT)),
                        "value_chain_supplier", "pt", COLOR_PT, POSITIONS + ["(blank)"]))
    with c2:
        chart_card("Value chain as buyer", "Customers by their position when buying.",
                   hbar(counts_by(pt_split, "value_chain_buyer", "pt", list(COLOR_PT)),
                        "value_chain_buyer", "pt", COLOR_PT, POSITIONS + ["(blank)"]))
    with c3:
        subs = [x if x else "(blank)" for x in (ordered(df["subsegment"]) or [""])]
        color_sub = {s_: PALETTE[i % len(PALETTE)] for i, s_ in enumerate(subs)}
        chart_card("Customers by region", "Split by subsegment.",
                   hbar(counts_by(fdf.assign(subsegment=fdf["subsegment"].replace("", "(blank)")), "region", "subsegment", subs),
                        "region", "subsegment", color_sub, legend_columns=0), legend=color_sub)
    chart_card("Customers by main sector", "Split by whether the customer owns a PT. Largest sector first.",
               hbar(counts_by(pt_split, "economicsector_segment", "pt", list(COLOR_PT)), "economicsector_segment", "pt", COLOR_PT))
    st.write("")

    # match results
    st.subheader("Match results")
    st.caption(f"Every customer in the filter with their top {TOP_PER_RELATION} partner candidates per relation: "
               "supplier (from their upstream column), buyer (downstream), two-way (midstream) and enabling. "
               f"Partners are other customers who own a PT in those sectors (whole portfolio), ranked by {prank_name} "
               "(sectors listed first get a slight preference). Change it with Rank partners by, under the filters.")
    st.caption(f"{len(matches):,} matches for {matches['customerid'].nunique() if not matches.empty else 0:,} "
               f"of {len(fdf):,} customers.")
    if not matches.empty:
        show = matches[["customer_rank", "customerid", "customername", "customercompany", "economicsector_segment", "relation",
                        "partner_rank", "partner_company", "partner_customername", "partner_sector",
                        "partner_region", "partner_sto_annum", "partner_bizvolume"]]
        st.dataframe(show, use_container_width=True, hide_index=True, height=360,
                     column_config={"customer_rank": st.column_config.NumberColumn("rank", help=f"Customer rank by {rank_name}"),
                                    "partner_rank": st.column_config.NumberColumn("partner_rank", help=f"Partner order by {prank_name}"),
                                    "partner_sto_annum": st.column_config.NumberColumn("partner turnover/yr", format="%.0f"),
                                    "partner_bizvolume": st.column_config.NumberColumn("partner biz volume", format="%.0f")})
    download("Download match results (CSV)", matches, "matchmaking_results", "dl_matches")

    # customer detail
    st.subheader("Customer detail")
    detail_pool = fdf if not fdf.empty else df
    detail_labels = detail_pool.sort_values(["has_pt_ocbc", rank_col], ascending=[False, False], na_position="last")["label"].tolist()
    pick = st.selectbox("Customer", detail_labels, key="detail")
    c = df.loc[df["label"] == pick].iloc[0]
    st.markdown(f"**Showing:** `{c['customerid']} · {c['customername']}`")
    lists, fallback = customer_chain(c, df)
    if fallback:
        st.caption("This customer's chain columns are empty, so the combined chain of their sector is used.")
    p1, p2, p3 = st.columns(3)
    n_pt_c = int(c["total_PT"]) if pd.notna(c["total_PT"]) else 0
    company = ("No PT" if c["has_pt_ocbc"] != "Y" else
               f"{c['customercompany'] or 'Has a PT (name not in data)'}" + (f" ({n_pt_c:,} PT)" if n_pt_c else ""))
    p1.markdown(f"**Company**  \n{company}")
    p1.markdown(f"**Subsegment**  \n{c['subsegment']}")
    p2.markdown(f"**Sector**  \n{c['economicsector_segment']}" + (f" · {c['product']}" if c.get('product') else ""))
    p2.markdown(f"**Value chain (supplier · buyer)**  \n{c['value_chain_supplier'] or '—'} · {c['value_chain_buyer'] or '—'}")
    p3.markdown(f"**Region · branch**  \n{c['region']} · {c['officename']}")
    p3.markdown(f"**Annual company turnover**  \n{rp(c['sto_annum']) if c['has_pt_ocbc'] == 'Y' else '—'}")
    p1.markdown(f"**Biz volume (AUM + LUM)**  \n{rp(c['bizvolume'])}")
    p2.markdown(f"**AUM · LUM**  \n{rp(c['aum_indiv'])} · {rp(c['lum_indiv'])}")
    in_f = fdf.loc[fdf["customerid"] == c["customerid"], "customer_rank"]
    p3.markdown(f"**Rank by {rank_name}**  \n" + (f"#{int(in_f.iloc[0]):,} of {len(fdf):,} in the current filter" if len(in_f) else "outside the current filter"))

    st.markdown("**This customer's sector chain** (number = customers with a PT in that sector)")
    cc = st.columns(4)
    for col, key in zip(cc, ["up", "mid", "down", "en"]):
        col.markdown(f"*{POS_LABEL[key]}*")
        col.markdown("  \n".join(f"`{s}` {int(pt_counts_all.get(s, 0)):,}" for s in lists[key]) or "None")

    pool = pt_pool(data_key, df, prank)
    st.caption(f"Partner candidates below are ranked by {prank_name}.")
    cand_tabs = st.tabs([REL_LABEL[k] for k in ["up", "down", "mid", "en"]])
    for t, key in zip(cand_tabs, ["up", "down", "mid", "en"]):
        with t:
            cand = candidates(c, key, TOP_IN_DETAIL, df, pool, lists, prank)
            if not cand:
                st.caption("No customers with a PT in these sectors yet." if lists[key] else "No sectors listed.")
            else:
                st.dataframe(pd.DataFrame([{
                    "company": x["r"]["customercompany"], "owner": x["r"]["customername"],
                    "CIF": x["r"]["customerid"], "sector": x["r"]["economicsector_segment"],
                    "region": x["r"]["region"], "turnover/yr": x["r"]["sto_annum"],
                    "biz volume": x["r"]["bizvolume"], "AUM": x["r"]["aum_indiv"], "LUM": x["r"]["lum_indiv"],
                } for x in cand]), use_container_width=True, hide_index=True,
                    column_config={c_: st.column_config.NumberColumn(format="%.0f") for c_ in ["turnover/yr", "biz volume", "AUM", "LUM"]})

    # sector pairs
    st.subheader("Top supplier → buyer sector pairs")
    st.caption("Built from the filtered customers: a pair counts when one of them in the supplier sector lists the "
               "buyer sector as downstream, or one in the buyer sector lists the supplier sector as upstream. "
               "Mentions = filtered customers listing the link. PT owners are counted over the whole portfolio, "
               "since partners can come from anywhere. Combinations = supplier PT owners × buyer PT owners.")
    pairs = sector_pairs(data_key, focus_ids, df)
    st.dataframe(pairs.head(20), use_container_width=True, hide_index=True)
    download("Download sector pairs (CSV)", pairs, "sector_pairs", "dl_pairs")

    # portfolio
    st.subheader("Customer portfolio")
    view_cols = ["customer_rank", "customerid", "customername", "subsegment", "customercompany", "economicsector_segment",
                 "product", "value_chain_supplier", "value_chain_buyer", "officename",
                 "has_pt_ocbc", "bizvolume", "aum_indiv", "lum_indiv", "sto", "avrg_trx_3m", "total_PT"]
    ranked = fdf.sort_values("customer_rank")
    st.caption(f"Sorted by {rank_name}, biggest first.")
    st.dataframe(ranked[[c_ for c_ in view_cols if c_ in fdf.columns]],
                 use_container_width=True, hide_index=True, height=360,
                 column_config={"customer_rank": st.column_config.NumberColumn("rank"),
                                "bizvolume": st.column_config.NumberColumn("biz volume", format="%.0f"),
                                "aum_indiv": st.column_config.NumberColumn("AUM", format="%.0f"),
                                "lum_indiv": st.column_config.NumberColumn("LUM", format="%.0f"),
                                "sto": st.column_config.NumberColumn("turnover/mo", format="%.0f")})
    port = raw.loc[ranked.index, raw_columns].copy()
    port.insert(0, "customer_rank", ranked["customer_rank"])
    if "bizvolume" not in resolve_columns(raw_columns):
        port["bizvolume"] = ranked["bizvolume"]
    download("Download portfolio (CSV, all columns)", port, "customer_portfolio", "dl_portfolio")

# ------------------------------------------------------------ Sector chain tab
with tab_s:
    st.caption("Explore how sectors connect, for the whole portfolio. This page has its own sector picker: "
               "the filters on the Matchmaking tab don't apply here.")
    linked = {x for col in ["list_up", "list_mid", "list_down", "list_en"] for lst in df[col] for x in lst}
    present = sorted(set(df["economicsector_segment"]) - {"", "(no sector)"}) + sorted(linked - set(df["economicsector_segment"]))
    default_sector = [c["economicsector_segment"]] if c["economicsector_segment"] in present else []
    secs = st.multiselect("Sectors", present, default=default_sector, key="chain_sectors",
                          help="Each selected sector gets its own chain map.")
    box = st.container(border=True)
    box.markdown(f"**Showing {len(secs)} sector{'s' if len(secs) != 1 else ''}**" + ("" if secs else " · pick a sector above"))
    if secs:
        box.markdown(chips("Sectors", secs))
    st.caption("Each sector's chain is combined from the upstream, midstream, downstream and enabling columns of "
               "every customer in that sector (whole portfolio). Arrows follow the flow of goods; thicker arrows = "
               "listed by more of the sector's customers. When customers place a sector in different columns, it "
               "goes where it is listed most.")
    st.markdown(
        f"<span style='color:{COLOR_POS['up']}'>━━</span> Upstream: supplies the selected sector &nbsp; "
        f"<span style='color:{COLOR_POS['mid']}'>━━</span> Midstream: two-way &nbsp; "
        f"<span style='color:{COLOR_POS['down']}'>━━</span> Downstream: buys from the selected sector",
        unsafe_allow_html=True)

    detail_rows = []
    for sec in secs:
        n, links = chain_of(sec, df)
        st.markdown(f"### {sec}")
        if n == 0:
            st.caption("No customers in this sector yet, so there is no chain to draw.")
            continue
        prod = links[(links["pos"] != "en") & (links["linked_sector"] != sec)]
        if prod.empty:
            st.caption("Customers in this sector have no upstream, midstream or downstream sectors filled in.")
        else:
            st.graphviz_chart(chain_dot(sec, n, links, pt_counts_all), use_container_width=True)
        en = links[(links["pos"] == "en") & (links["linked_sector"] != sec)]
        st.markdown("**Enabling sectors**")
        if en.empty:
            st.caption("No enabling sectors listed.")
        else:
            ecols = st.columns(min(4, len(en)))
            for i, (_, l) in enumerate(en.iterrows()):
                with ecols[i % len(ecols)]:
                    st.markdown(f"**{l['linked_sector']}**  \nlisted by {l['customers_listing']:,} of {n:,} "
                                f"({l['share']:.0%}) · {int(pt_counts_all.get(l['linked_sector'], 0)):,} with a PT")
                    st.progress(min(1.0, float(l["share"])))
        for _, l in links[links["linked_sector"] != sec].iterrows():
            detail_rows.append({"sector": sec, "sector_customers": n, "linked_sector": l["linked_sector"],
                                "position": POS_LABEL[l["pos"]], "customers_listing": l["customers_listing"],
                                "share_of_sector_customers": round(l["share"], 4),
                                "linked_sector_pt_owners": int(pt_counts_all.get(l["linked_sector"], 0))})
    st.subheader("Chain details")
    chain_table = pd.DataFrame(detail_rows)
    if not chain_table.empty:
        order = {"Upstream": 0, "Midstream": 1, "Downstream": 2, "Enabling": 3}
        chain_table = chain_table.sort_values(["sector", "position", "customers_listing"],
                                              key=lambda s: s.map(order) if s.name == "position" else s,
                                              ascending=[True, True, False])
        st.dataframe(chain_table, use_container_width=True, hide_index=True,
                     column_config={"share_of_sector_customers": st.column_config.ProgressColumn("share of sector customers", format="%.2f", min_value=0, max_value=1)})
    download("Download chain details (CSV)", chain_table, "sector_chain", "dl_chain")


# ------------------------------------------------------------ Unmapped data tab
MISSING_CHECKS = {
    "Sector chain (Upstream, Mid-stream, Downstream, Enabling all empty)": "no_chain",
    "Value chain as supplier": "no_vc_supplier",
    "Value chain as buyer": "no_vc_buyer",
    "Main sector": "no_sector",
}
MISSING_SHORT = {"no_chain": "sector chain", "no_vc_supplier": "value chain supplier",
                 "no_vc_buyer": "value chain buyer", "no_sector": "main sector"}

with tab_u:
    st.caption("Customers whose mapping columns are empty in the file, for the whole portfolio (the Matchmaking filters "
               "don't apply here). Fill these in the Excel file and upload it again to map them.")
    u1, u2, u3, u4 = st.columns(4)
    u1.metric("No sector chain", f"{int(df['no_chain'].sum()):,}")
    u2.metric("No value chain (supplier)", f"{int(df['no_vc_supplier'].sum()):,}")
    u3.metric("No value chain (buyer)", f"{int(df['no_vc_buyer'].sum()):,}")
    u4.metric("No main sector", f"{int(df['no_sector'].sum()):,}")
    st.caption("A customer with no sector chain still gets partner suggestions from the combined chain of other customers "
               "in the same sector. If nobody in that sector is mapped, they get none.")

    checks = st.multiselect("Show customers missing any of", list(MISSING_CHECKS), default=list(MISSING_CHECKS)[:3], key="unmapped_checks")
    cols_u = [MISSING_CHECKS[c] for c in checks]
    if not cols_u:
        st.info("Pick at least one item above.")
    else:
        um = df[df[cols_u].any(axis=1)].copy()
        um["missing"] = um[cols_u].apply(lambda r: ", ".join(MISSING_SHORT[c] for c in cols_u if r[c]), axis=1)
        sectors_u = ordered(um["economicsector_segment"])
        pick_sec = st.multiselect("Main sector", sectors_u, key="unmapped_sectors", placeholder="All sectors")
        if pick_sec:
            um = um[um["economicsector_segment"].isin(pick_sec)]
        st.markdown(f"**{len(um):,} customers** missing at least one of the picked items "
                    f"({len(um) / len(df):.1%} of the portfolio).")

        st.subheader("By sector and sub-sector")
        st.caption("Each row is a sub-sector that needs mapping. Rows with the most customers first: mapping those "
                   "covers the most customers.")
        grp_cols = [c for c in ["economicsector_segment", "econ_subsegment", "product"] if um[c].ne("").any()]
        if um.empty:
            st.success("Nothing unmapped for these items.")
        else:
            agg = {"customers": ("customerid", "size")}
            for c in cols_u:
                agg[MISSING_SHORT[c] + " empty"] = (c, "sum")
            by = um.groupby(grp_cols, dropna=False).agg(**agg).reset_index().sort_values("customers", ascending=False)
            by = by.rename(columns={"economicsector_segment": "main sector", "econ_subsegment": "economicsector_subsegment",
                                    "product": "product / sub-sector"})
            st.dataframe(by, use_container_width=True, hide_index=True, height=320)
            download("Download sub-sectors to map (CSV)", by, "unmapped_subsectors", "dl_unmapped_groups")

            st.subheader("Customer list")
            list_cols = [c for c in ["customerid", "customername", "missing", "subsegment", "region", "area", "officename",
                                     "economicsector_segment", "econ_bicode", "econ_subsegment", "product",
                                     "value_chain_supplier", "value_chain_buyer", "has_pt_ocbc"]
                         if c in um.columns and (c in ("missing", "customerid") or um[c].ne("").any())]
            st.dataframe(um[list_cols].rename(columns={"econ_bicode": "economicsector_bicode",
                                                       "econ_subsegment": "economicsector_subsegment",
                                                       "economicsector_segment": "main sector"}),
                         use_container_width=True, hide_index=True, height=360)
            out_u = raw.loc[um.index, raw_columns].copy()
            out_u.insert(0, "missing", um["missing"])
            download("Download unmapped customers (CSV, all columns)", out_u, "unmapped_customers", "dl_unmapped")
