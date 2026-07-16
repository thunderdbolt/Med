from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DEFAULT_DATA = APP_DIR / "data" / "Emerging_Medical_Devices_Research.xlsx"
IMAGE_DIR = APP_DIR / "images"
IMAGE_MANIFEST = APP_DIR / "data" / "image_manifest.csv"

st.set_page_config(
    page_title="Emerging Medical Device Explorer",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      :root {
        --ink: #183153;
        --muted: #64748b;
        --teal: #0f766e;
        --teal-soft: #ecfdf5;
        --blue-soft: #eff6ff;
        --amber-soft: #fffbeb;
        --red-soft: #fef2f2;
        --line: #e2e8f0;
        --surface: #ffffff;
      }
      .stApp {background: linear-gradient(180deg, #f8fbff 0%, #ffffff 32%);}
      .block-container {padding-top: 1.35rem; padding-bottom: 4rem; max-width: 1450px;}
      [data-testid="stSidebar"] {border-right: 1px solid var(--line);}
      [data-testid="stMetric"] {
        background: rgba(255,255,255,.92); border: 1px solid var(--line);
        padding: 16px 18px; border-radius: 18px; box-shadow: 0 8px 24px rgba(15,23,42,.05);
      }
      .hero {
        border: 1px solid #dbeafe; border-radius: 26px; padding: 30px 32px;
        background: radial-gradient(circle at 86% 18%, rgba(20,184,166,.16), transparent 25%),
                    linear-gradient(120deg, #eff6ff 0%, #f8fafc 55%, #ecfeff 100%);
        box-shadow: 0 16px 45px rgba(15,23,42,.07); margin-bottom: 18px;
      }
      .eyebrow {font-size:.78rem; letter-spacing:.12em; text-transform:uppercase; font-weight:800; color:var(--teal);}
      .hero h1 {color:var(--ink); font-size:2.45rem; line-height:1.08; margin:.45rem 0 .7rem;}
      .hero p {color:var(--muted); max-width:850px; font-size:1.03rem; margin:0;}
      .product-card {
        background: rgba(255,255,255,.97); border: 1px solid var(--line); border-radius: 22px;
        padding: 16px; margin-bottom: 16px; box-shadow: 0 10px 28px rgba(15,23,42,.055);
      }
      .product-title {font-size:1.22rem; line-height:1.25; font-weight:800; color:var(--ink); margin:.15rem 0 .35rem;}
      .muted {color:var(--muted); font-size:.92rem;}
      .badge {display:inline-block; padding:5px 10px; margin:5px 5px 2px 0; border-radius:999px; font-size:.76rem; font-weight:750; border:1px solid transparent;}
      .badge-good {background:var(--teal-soft); color:#047857; border-color:#a7f3d0;}
      .badge-warn {background:var(--amber-soft); color:#a16207; border-color:#fde68a;}
      .badge-info {background:var(--blue-soft); color:#1d4ed8; border-color:#bfdbfe;}
      .badge-muted {background:#f8fafc; color:#475569; border-color:#e2e8f0;}
      .section-label {font-size:.76rem; font-weight:800; letter-spacing:.08em; text-transform:uppercase; color:#64748b; margin:.7rem 0 .2rem;}
      .detail-shell {background:white; border:1px solid var(--line); border-radius:24px; padding:22px; box-shadow:0 14px 34px rgba(15,23,42,.06);}
      .profile-title {color:var(--ink); font-weight:850; font-size:2rem; line-height:1.15; margin:.25rem 0 .6rem;}
      .info-panel {background:#f8fafc; border:1px solid var(--line); border-radius:16px; padding:15px 17px; min-height:120px;}
      .info-panel h4 {font-size:.78rem; text-transform:uppercase; letter-spacing:.07em; color:#64748b; margin:0 0 .55rem;}
      .info-panel p {margin:0; color:#334155; line-height:1.55;}
      .stat-strip {background:white; border:1px solid var(--line); border-radius:18px; padding:15px; text-align:center;}
      .stat-number {font-size:1.8rem; font-weight:850; color:var(--ink);}
      .stat-label {font-size:.82rem; color:var(--muted);}
      div.stButton > button {border-radius:12px; font-weight:700;}
      div[data-testid="stImage"] img {border-radius:16px;}
    </style>
    """,
    unsafe_allow_html=True,
)

COLUMN_ALIASES = {
    "manufacturer/ developer": "Manufacturer/Developer",
    "manufacturer/developer": "Manufacturer/Developer",
    "u.s status": "U.S. Status",
    "u.s. status": "U.S. Status",
    "europe status": "Europe Status",
    "middle east status": "Middle East Status",
    "commercial status": "Commercial Status",
    "product type": "Product Type",
    "country of origin": "Country of Origin",
    "medical field": "Medical Field",
    "hospital need": "Hospital Need",
    "material used": "Material Used",
    "market potential": "Market Potential",
    "disadvantages/limitations": "Disadvantages/Limitations",
}

REQUIRED_COLUMNS = ["Product Name", "Product Type", "Country of Origin"]
TEXT_COLUMNS = [
    "Product Name", "Product Type", "Country of Origin", "Website",
    "Manufacturer/Developer", "Medical Field", "Hospital Need", "Innovation",
    "Material Used", "Advantages", "Disadvantages/Limitations", "U.S. Status",
    "Europe Status", "Middle East Status", "Commercial Status", "Market Potential",
    "Recommendation", "Image Path", "Gallery Paths",
]


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def split_bullets(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    parts = [p.strip(" •-\t") for p in text.replace("\n", " • ").split("•")]
    return [p for p in parts if p and p not in {"...", "[…]", "[ ... ]"}]


def first_nonempty(items: Iterable[str], fallback: str = "Not specified") -> str:
    return next((item for item in items if item), fallback)


def normalize_status(value: object) -> str:
    text = clean_text(value).lower()
    if not text:
        return "Not stated"
    negative = ("no public", "not confirmed", "not identified", "pending", "awaiting", "underway")
    if any(term in text for term in ("approved", "clearance", "cleared", "certification reported", "ce certified")):
        return "Unconfirmed / pending" if any(term in text for term in negative) else "Approved / certified"
    if any(term in text for term in ("commercial", "available", "marketed", "sales")):
        return "Limited / unconfirmed" if any(term in text for term in negative) else "Commercially available"
    if any(term in text for term in ("development", "early-stage", "pre-commercial", "planned")):
        return "Under development"
    if any(term in text for term in negative):
        return "Unconfirmed / pending"
    return "Other"


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {col: COLUMN_ALIASES.get(clean_text(col).lower(), clean_text(col)) for col in df.columns}
    df = df.rename(columns=renamed)
    for col in TEXT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(clean_text)
    return df


@st.cache_data(show_spinner=False)
def load_image_manifest(path: str) -> pd.DataFrame:
    manifest_path = Path(path)
    if not manifest_path.exists():
        return pd.DataFrame(columns=["No.", "Product Name", "Image Path", "Gallery Paths"])
    manifest = pd.read_csv(manifest_path)
    manifest["No."] = pd.to_numeric(manifest.get("No."), errors="coerce")
    return manifest


def attach_image_manifest(df: pd.DataFrame) -> pd.DataFrame:
    manifest = load_image_manifest(str(IMAGE_MANIFEST))
    if manifest.empty:
        return df
    result = df.copy()
    if "No." in result.columns:
        result["No."] = pd.to_numeric(result["No."], errors="coerce")
        result = result.merge(manifest[["No.", "Image Path", "Gallery Paths"]], on="No.", how="left", suffixes=("", "_manifest"))
    else:
        result = result.merge(manifest[["Product Name", "Image Path", "Gallery Paths"]], on="Product Name", how="left", suffixes=("", "_manifest"))
    for column in ("Image Path", "Gallery Paths"):
        manifest_column = f"{column}_manifest"
        if manifest_column in result.columns:
            result[column] = result[column].fillna("").astype(str)
            result[column] = result[column].where(result[column].str.strip().ne(""), result[manifest_column].fillna(""))
            result = result.drop(columns=[manifest_column])
    return result


def detect_excel_header(source) -> int:
    preview = pd.read_excel(source, header=None, nrows=12)
    for idx, row in preview.iterrows():
        values = {clean_text(v).lower() for v in row.tolist()}
        if "product name" in values and ("country of origin" in values or "product type" in values):
            return int(idx)
    return 0


def read_data(source) -> pd.DataFrame:
    name = getattr(source, "name", str(source)).lower()
    if name.endswith(".csv"):
        df = pd.read_csv(source)
    else:
        header_row = detect_excel_header(source)
        if hasattr(source, "seek"):
            source.seek(0)
        df = pd.read_excel(source, header=header_row)
    df = normalize_columns(df)
    df = df[df["Product Name"].str.len() > 0].copy()
    placeholder = df["Product Name"].str.fullmatch(r"\[?\s*\.{2,}\s*\]?", case=False, na=False)
    df = df[~placeholder].copy()
    df["US Category"] = df["U.S. Status"].map(normalize_status)
    df["Europe Category"] = df["Europe Status"].map(normalize_status)
    df["Middle East Category"] = df["Middle East Status"].map(normalize_status)
    df["Commercial Category"] = df["Commercial Status"].map(normalize_status)
    df["Primary Medical Field"] = df["Medical Field"].map(lambda x: first_nonempty(split_bullets(x)))
    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_default_data(path: str) -> pd.DataFrame:
    return attach_image_manifest(read_data(Path(path)))


def validate_data(df: pd.DataFrame) -> list[str]:
    return [col for col in REQUIRED_COLUMNS if col not in df.columns or df[col].eq("").all()]


def options_from_bullets(series: pd.Series) -> list[str]:
    values: set[str] = set()
    for value in series:
        values.update(split_bullets(value))
    return sorted(values)


def contains_any(value: str, selected: list[str]) -> bool:
    return not selected or any(item.lower() in value.lower() for item in selected)


def website_url(value: str) -> str | None:
    candidate = first_nonempty(split_bullets(value), "")
    if not candidate or "no dedicated" in candidate.lower():
        return None
    return candidate if candidate.startswith(("http://", "https://")) else f"https://{candidate.strip()}"


def resolve_gallery(row: pd.Series) -> list[Path]:
    raw_paths = clean_text(row.get("Gallery Paths", "")) or clean_text(row.get("Image Path", ""))
    resolved: list[Path] = []
    for raw_path in [part.strip() for part in raw_paths.split("|") if part.strip()]:
        candidates = [APP_DIR / raw_path, IMAGE_DIR / Path(raw_path).name]
        match = next((path for path in candidates if path.exists()), None)
        if match and match not in resolved:
            resolved.append(match)
    return resolved


def resolve_image(row: pd.Series) -> Path | None:
    gallery = resolve_gallery(row)
    return gallery[0] if gallery else None


def badge_class(status: str) -> str:
    lowered = status.lower()
    if "approved" in lowered or "commercially" in lowered:
        return "badge-good"
    if "pending" in lowered or "development" in lowered or "limited" in lowered:
        return "badge-warn"
    if "not stated" in lowered or "other" in lowered:
        return "badge-muted"
    return "badge-info"


def badge_html(label: str, status: str) -> str:
    return f'<span class="badge {badge_class(status)}">{escape(label)}: {escape(status)}</span>'


def short_text(value: str, max_chars: int = 190) -> str:
    text = first_nonempty(split_bullets(value), "Not stated")
    return text if len(text) <= max_chars else text[: max_chars - 1].rstrip() + "…"


def render_profile(row: pd.Series) -> None:
    top_left, top_right = st.columns([1.15, 1.55], gap="large")
    gallery = resolve_gallery(row)
    with top_left:
        if gallery:
            index_key = f"gallery_index_{int(row.get('No.', 0) or 0)}"
            if index_key not in st.session_state:
                st.session_state[index_key] = 0
            index = min(st.session_state[index_key], len(gallery) - 1)
            st.image(str(gallery[index]), use_container_width=True)
            if len(gallery) > 1:
                prev_col, counter_col, next_col = st.columns([1, 2, 1])
                if prev_col.button("←", key=f"prev_{index_key}", use_container_width=True):
                    st.session_state[index_key] = (index - 1) % len(gallery)
                    st.rerun()
                counter_col.markdown(f"<div style='text-align:center;color:#64748b;padding:.45rem'>{index + 1} of {len(gallery)}</div>", unsafe_allow_html=True)
                if next_col.button("→", key=f"next_{index_key}", use_container_width=True):
                    st.session_state[index_key] = (index + 1) % len(gallery)
                    st.rerun()
                thumbs = st.columns(min(4, len(gallery)))
                for i, image in enumerate(gallery):
                    with thumbs[i % len(thumbs)]:
                        st.image(str(image), use_container_width=True)
        else:
            st.info("No product image has been linked yet.")

    with top_right:
        st.markdown('<div class="eyebrow">Product intelligence profile</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="profile-title">{escape(row["Product Name"])}</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="muted">{escape(row["Country of Origin"])} · {escape(first_nonempty(split_bullets(row["Product Type"])))}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            badge_html("Commercial", row["Commercial Category"])
            + badge_html("U.S.", row["US Category"])
            + badge_html("Europe", row["Europe Category"])
            + badge_html("Middle East", row["Middle East Category"]),
            unsafe_allow_html=True,
        )
        st.markdown('<div class="section-label">Innovation snapshot</div>', unsafe_allow_html=True)
        st.write(short_text(row["Innovation"], 500))
        meta1, meta2 = st.columns(2)
        with meta1:
            st.markdown("**Manufacturer / developer**")
            st.write(row["Manufacturer/Developer"] or "Not stated")
        with meta2:
            st.markdown("**Medical field**")
            st.write(row["Medical Field"] or "Not stated")
        url = website_url(row["Website"])
        if url:
            st.link_button("Visit product or manufacturer website ↗", url, use_container_width=True)

    st.markdown("### Research profile")
    overview_tab, clinical_tab, market_tab, regulatory_tab = st.tabs(["Overview", "Clinical value", "Market", "Regulatory"])
    with overview_tab:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f'<div class="info-panel"><h4>Innovation</h4><p>{escape(row["Innovation"] or "Not stated")}</p></div>', unsafe_allow_html=True)
            st.write("")
            st.markdown(f'<div class="info-panel"><h4>Materials</h4><p>{escape(row["Material Used"] or "Not stated")}</p></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="info-panel"><h4>Product type</h4><p>{escape(row["Product Type"] or "Not stated")}</p></div>', unsafe_allow_html=True)
            st.write("")
            st.markdown(f'<div class="info-panel"><h4>Medical field</h4><p>{escape(row["Medical Field"] or "Not stated")}</p></div>', unsafe_allow_html=True)
    with clinical_tab:
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="info-panel"><h4>Hospital need</h4><p>{escape(row["Hospital Need"] or "Not stated")}</p></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="info-panel"><h4>Advantages</h4><p>{escape(row["Advantages"] or "Not stated")}</p></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="info-panel"><h4>Limitations</h4><p>{escape(row["Disadvantages/Limitations"] or "Not stated")}</p></div>', unsafe_allow_html=True)
    with market_tab:
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="info-panel"><h4>Commercial status</h4><p>{escape(row["Commercial Status"] or "Not stated")}</p></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="info-panel"><h4>Market potential</h4><p>{escape(row["Market Potential"] or "Not stated")}</p></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="info-panel"><h4>Recommendation</h4><p>{escape(row["Recommendation"] or "Not stated")}</p></div>', unsafe_allow_html=True)
    with regulatory_tab:
        for label, text in (("United States", row["U.S. Status"]), ("Europe", row["Europe Status"]), ("Middle East", row["Middle East Status"])):
            st.markdown(f"**{label}**")
            st.write(text or "Not stated")
            st.divider()


# Session defaults
st.session_state.setdefault("selected_product", None)
st.session_state.setdefault("page", "Dashboard")

with st.sidebar:
    st.markdown("## 🩺 Device Explorer")
    st.caption("Emerging medical technology intelligence")
    page = st.radio("Navigate", ["Dashboard", "Explorer", "Compare", "Analytics", "Data"], key="page")
    st.divider()
    st.markdown("### Data source")
    upload = st.file_uploader("Upload Excel or CSV", type=["xlsx", "xls", "csv"], help="Used only for this browser session.")
    try:
        data = attach_image_manifest(read_data(upload)) if upload else load_default_data(str(DEFAULT_DATA))
    except Exception as exc:
        st.error(f"Could not read the data file: {exc}")
        st.stop()
    missing = validate_data(data)
    if missing:
        st.error("Missing required columns: " + ", ".join(missing))
        st.stop()
    st.success(f"{len(data)} valid products loaded")

    st.divider()
    st.markdown("### Filters")
    search = st.text_input("Search", placeholder="Product, material, innovation…")
    countries = st.multiselect("Country", sorted(data["Country of Origin"].dropna().unique()))
    product_types = st.multiselect("Product type", options_from_bullets(data["Product Type"]))
    fields = st.multiselect("Medical field", options_from_bullets(data["Medical Field"]))
    commercial = st.multiselect("Commercial status", sorted(data["Commercial Category"].unique()))
    us_status = st.multiselect("U.S. status", sorted(data["US Category"].unique()))
    eu_status = st.multiselect("Europe status", sorted(data["Europe Category"].unique()))

filtered = data.copy()
if search:
    searchable_cols = [c for c in TEXT_COLUMNS if c in filtered.columns]
    mask = filtered[searchable_cols].apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
    filtered = filtered[mask]
if countries:
    filtered = filtered[filtered["Country of Origin"].isin(countries)]
if product_types:
    filtered = filtered[filtered["Product Type"].map(lambda x: contains_any(x, product_types))]
if fields:
    filtered = filtered[filtered["Medical Field"].map(lambda x: contains_any(x, fields))]
if commercial:
    filtered = filtered[filtered["Commercial Category"].isin(commercial)]
if us_status:
    filtered = filtered[filtered["US Category"].isin(us_status)]
if eu_status:
    filtered = filtered[filtered["Europe Category"].isin(eu_status)]

# Dedicated product profile view
if st.session_state.selected_product:
    product_rows = data[data["Product Name"] == st.session_state.selected_product]
    if product_rows.empty:
        st.session_state.selected_product = None
        st.rerun()
    back_col, spacer = st.columns([1, 5])
    if back_col.button("← Back to explorer", use_container_width=True):
        st.session_state.selected_product = None
        st.session_state.page = "Explorer"
        st.rerun()
    st.markdown('<div class="detail-shell">', unsafe_allow_html=True)
    render_profile(product_rows.iloc[0])
    st.markdown('</div>', unsafe_allow_html=True)
    st.divider()
    st.caption("Research support tool only. Verify regulatory and commercial information before procurement or clinical decisions.")
    st.stop()

if page == "Dashboard":
    st.markdown(
        """
        <div class="hero">
          <div class="eyebrow">Emerging medical technology intelligence</div>
          <h1>Explore promising devices, markets, and regulatory pathways.</h1>
          <p>Filter the research portfolio, inspect rich product profiles, compare technologies, and explore geographic and commercial patterns.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Products", len(filtered))
    m2.metric("Countries", filtered["Country of Origin"].nunique())
    m3.metric("Medical fields", filtered["Primary Medical Field"].nunique())
    m4.metric("Commercially available", int((filtered["Commercial Category"] == "Commercially available").sum()))

    st.markdown("### Global portfolio")
    if filtered.empty:
        st.info("No products match the current filters.")
    else:
        map_df = filtered.groupby("Country of Origin", as_index=False).agg(Products=("Product Name", "count"), Product_names=("Product Name", lambda s: "<br>".join(s)))
        fig = px.scatter_geo(map_df, locations="Country of Origin", locationmode="country names", size="Products", hover_name="Country of Origin", hover_data={"Products": True, "Product_names": True}, projection="natural earth")
        fig.update_geos(showframe=False, showcoastlines=True, coastlinecolor="#cbd5e1", landcolor="#f8fafc", bgcolor="rgba(0,0,0,0)")
        fig.update_layout(height=470, margin=dict(l=0, r=0, t=10, b=0), paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Featured products")
    featured = filtered.head(3)
    cols = st.columns(3)
    for idx, (_, row) in enumerate(featured.iterrows()):
        with cols[idx]:
            st.markdown('<div class="product-card">', unsafe_allow_html=True)
            image = resolve_image(row)
            if image:
                st.image(str(image), use_container_width=True)
            st.markdown(f'<div class="product-title">{escape(row["Product Name"])}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="muted">{escape(row["Country of Origin"])} · {escape(row["Primary Medical Field"])}</div>', unsafe_allow_html=True)
            st.markdown(badge_html("Commercial", row["Commercial Category"]), unsafe_allow_html=True)
            st.write(short_text(row["Innovation"], 145))
            if st.button("View full profile", key=f"dash_{row['Product Name']}", use_container_width=True):
                st.session_state.selected_product = row["Product Name"]
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

elif page == "Explorer":
    st.markdown("## Product explorer")
    st.caption("Browse the portfolio using the filters in the sidebar.")
    top_a, top_b = st.columns([3, 1])
    with top_a:
        st.write(f"Showing **{len(filtered)}** of **{len(data)}** products")
    with top_b:
        sort_choice = st.selectbox("Sort", ["Product name", "Country", "Commercial status"], label_visibility="collapsed")
    sort_map = {"Product name": "Product Name", "Country": "Country of Origin", "Commercial status": "Commercial Category"}
    if filtered.empty:
        st.info("No products match the current filters.")
    else:
        rows = list(filtered.sort_values(sort_map[sort_choice]).iterrows())
        for start in range(0, len(rows), 3):
            cols = st.columns(3)
            for col, (_, row) in zip(cols, rows[start:start + 3]):
                with col:
                    st.markdown('<div class="product-card">', unsafe_allow_html=True)
                    image = resolve_image(row)
                    if image:
                        st.image(str(image), use_container_width=True)
                    else:
                        st.info("Image not linked")
                    st.markdown(f'<div class="product-title">{escape(row["Product Name"])}</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="muted">{escape(row["Country of Origin"])} · {escape(row["Primary Medical Field"])}</div>', unsafe_allow_html=True)
                    st.markdown(
                        badge_html("Commercial", row["Commercial Category"]) + badge_html("U.S.", row["US Category"]),
                        unsafe_allow_html=True,
                    )
                    st.markdown('<div class="section-label">Innovation</div>', unsafe_allow_html=True)
                    st.write(short_text(row["Innovation"]))
                    if st.button("View full profile", key=f"explore_{row['Product Name']}", use_container_width=True):
                        st.session_state.selected_product = row["Product Name"]
                        st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

elif page == "Compare":
    st.markdown("## Compare products")
    selected = st.multiselect("Choose up to four products", filtered["Product Name"].tolist(), max_selections=4)
    if not selected:
        st.info("Select products to create a side-by-side comparison.")
    else:
        selected_df = filtered[filtered["Product Name"].isin(selected)]
        cards = st.columns(len(selected))
        for col, (_, row) in zip(cards, selected_df.iterrows()):
            with col:
                image = resolve_image(row)
                if image:
                    st.image(str(image), use_container_width=True)
                st.markdown(f"**{row['Product Name']}**")
                st.caption(f"{row['Country of Origin']} · {row['Primary Medical Field']}")
        compare_cols = ["Product Name", "Product Type", "Country of Origin", "Medical Field", "Innovation", "Material Used", "Advantages", "Disadvantages/Limitations", "Commercial Category", "US Category", "Europe Category", "Middle East Category", "Recommendation"]
        comparison = selected_df[compare_cols].set_index("Product Name").T
        st.dataframe(comparison, use_container_width=True)

elif page == "Analytics":
    st.markdown("## Portfolio analytics")
    if filtered.empty:
        st.info("No records available for charts.")
    else:
        chart1, chart2 = st.columns(2)
        with chart1:
            country_counts = filtered["Country of Origin"].value_counts().rename_axis("Country").reset_index(name="Products")
            fig_country = px.bar(country_counts, x="Country", y="Products", title="Products by country")
            fig_country.update_layout(margin=dict(l=10, r=10, t=55, b=10))
            st.plotly_chart(fig_country, use_container_width=True)
        with chart2:
            status_counts = filtered["Commercial Category"].value_counts().rename_axis("Status").reset_index(name="Products")
            fig_status = px.pie(status_counts, names="Status", values="Products", hole=.52, title="Commercial readiness")
            st.plotly_chart(fig_status, use_container_width=True)
        field_counts = filtered["Primary Medical Field"].value_counts().head(12).rename_axis("Medical field").reset_index(name="Products")
        fig_field = px.bar(field_counts, x="Products", y="Medical field", orientation="h", title="Leading medical fields")
        st.plotly_chart(fig_field, use_container_width=True)

else:
    st.markdown("## Research data")
    display_cols = [c for c in ["Product Name", "Product Type", "Country of Origin", "Medical Field", "Commercial Category", "US Category", "Europe Category", "Middle East Category", "Website"] if c in filtered.columns]
    st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)
    csv = filtered.to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered data as CSV", csv, "filtered_medical_devices.csv", "text/csv")

st.divider()
st.caption("Research support tool only. Regulatory and commercial information should be independently verified before procurement or clinical decisions.")
