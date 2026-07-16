from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import streamlit as st

# -----------------------------------------------------------------------------
# App paths and page configuration
# -----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA = BASE_DIR / "data" / "Emerging_Medical_Devices_Research.xlsx"
IMAGE_DIR = BASE_DIR / "images"
IMAGE_MANIFEST = BASE_DIR / "data" / "image_manifest.csv"

st.set_page_config(
    page_title="Emerging Medical Device Explorer",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Visual system
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    :root {
        --navy: #10284c;
        --blue: #2463eb;
        --teal: #0f8b8d;
        --green: #16835b;
        --amber: #a16207;
        --red: #b42318;
        --ink: #172033;
        --muted: #667085;
        --line: #dbe2ea;
        --surface: #ffffff;
        --surface-soft: #f6f8fb;
    }

    .stApp {
        background:
            radial-gradient(circle at 88% 4%, rgba(36, 99, 235, .08), transparent 24rem),
            linear-gradient(180deg, #f7faff 0%, #ffffff 30%);
    }

    .block-container {
        max-width: 1420px;
        padding-top: 1.5rem;
        padding-bottom: 4rem;
    }

    [data-testid="stSidebar"] {
        border-right: 1px solid var(--line);
        min-width: 285px;
        max-width: 315px;
    }

    [data-testid="stSidebar"] .block-container {
        padding-top: 1.2rem;
    }

    [data-testid="stMetric"] {
        background: rgba(255, 255, 255, .92);
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 1rem 1.05rem;
        box-shadow: 0 8px 24px rgba(16, 40, 76, .05);
        min-height: 112px;
    }

    [data-testid="stMetricLabel"] p {
        white-space: normal;
        line-height: 1.2;
    }

    [data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 20px;
        border-color: var(--line);
        background: rgba(255, 255, 255, .96);
        box-shadow: 0 10px 30px rgba(16, 40, 76, .05);
    }

    div[data-testid="stImage"] img {
        border-radius: 14px;
    }

    div.stButton > button,
    div.stLinkButton > a {
        border-radius: 11px;
        font-weight: 700;
    }

    .hero {
        border: 1px solid #cfe0ff;
        border-radius: 28px;
        padding: 2.2rem 2.35rem;
        background:
            radial-gradient(circle at 88% 18%, rgba(15, 139, 141, .18), transparent 24%),
            linear-gradient(120deg, #edf4ff 0%, #f8fbff 55%, #ecfeff 100%);
        box-shadow: 0 18px 48px rgba(16, 40, 76, .08);
        margin-bottom: 1.15rem;
    }

    .eyebrow {
        color: var(--teal);
        font-size: .75rem;
        font-weight: 850;
        letter-spacing: .12em;
        text-transform: uppercase;
    }

    .hero h1 {
        color: var(--navy);
        font-size: clamp(2rem, 4vw, 3.25rem);
        line-height: 1.03;
        margin: .55rem 0 .8rem;
        max-width: 900px;
    }

    .hero p {
        color: var(--muted);
        font-size: 1.03rem;
        line-height: 1.65;
        max-width: 900px;
        margin: 0;
    }

    .product-title {
        color: var(--navy);
        font-size: 1.22rem;
        font-weight: 850;
        line-height: 1.25;
        margin: .35rem 0 .25rem;
    }

    .product-meta {
        color: var(--muted);
        font-size: .9rem;
        line-height: 1.4;
        margin-bottom: .35rem;
    }

    .section-label {
        color: #667085;
        font-size: .72rem;
        font-weight: 850;
        letter-spacing: .09em;
        text-transform: uppercase;
        margin-top: .8rem;
    }

    .badge {
        display: inline-block;
        border: 1px solid transparent;
        border-radius: 999px;
        font-size: .73rem;
        font-weight: 800;
        margin: .3rem .32rem .2rem 0;
        padding: .28rem .58rem;
    }

    .badge-approved {
        color: #067647;
        background: #ecfdf3;
        border-color: #abefc6;
    }

    .badge-development {
        color: #93370d;
        background: #fffaeb;
        border-color: #fedf89;
    }

    .badge-certified {
        color: #175cd3;
        background: #eff8ff;
        border-color: #b2ddff;
    }

    .badge-unconfirmed {
        color: #475467;
        background: #f2f4f7;
        border-color: #d0d5dd;
    }

    .badge-negative {
        color: #b42318;
        background: #fef3f2;
        border-color: #fecdca;
    }

    .placeholder {
        min-height: 215px;
        border: 1px dashed #b8c3d1;
        border-radius: 14px;
        background: linear-gradient(145deg, #f8fafc, #eef3f8);
        display: flex;
        align-items: center;
        justify-content: center;
        text-align: center;
        color: #667085;
        padding: 1rem;
        margin-bottom: .55rem;
    }

    .placeholder strong {
        color: var(--navy);
        display: block;
        margin-bottom: .25rem;
    }

    .profile-title {
        color: var(--navy);
        font-size: clamp(2rem, 4vw, 3rem);
        font-weight: 900;
        line-height: 1.05;
        margin: .4rem 0 .55rem;
    }

    .info-panel {
        background: var(--surface-soft);
        border: 1px solid var(--line);
        border-radius: 16px;
        min-height: 145px;
        padding: 1rem 1.1rem;
    }

    .info-panel h4 {
        color: #667085;
        font-size: .74rem;
        letter-spacing: .08em;
        margin: 0 0 .55rem;
        text-transform: uppercase;
    }

    .info-panel p {
        color: #344054;
        line-height: 1.6;
        margin: 0;
    }

    .small-note {
        color: var(--muted);
        font-size: .82rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Data normalization
# -----------------------------------------------------------------------------
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
    "Product Name",
    "Product Type",
    "Country of Origin",
    "Website",
    "Manufacturer/Developer",
    "Medical Field",
    "Hospital Need",
    "Innovation",
    "Material Used",
    "Advantages",
    "Disadvantages/Limitations",
    "U.S. Status",
    "Europe Status",
    "Middle East Status",
    "Commercial Status",
    "Market Potential",
    "Recommendation",
    "Image Path",
    "Gallery Paths",
]


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def split_bullets(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    normalized = text.replace("\n", " • ").replace(";", " • ")
    parts = [part.strip(" •-\t") for part in normalized.split("•")]
    placeholders = {"...", "[…]", "[ ... ]", "[... ]", "[... ]"}
    return [part for part in parts if part and part not in placeholders]


def first_nonempty(items: Iterable[str], fallback: str = "Not specified") -> str:
    return next((item for item in items if item), fallback)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {
        column: COLUMN_ALIASES.get(clean_text(column).lower(), clean_text(column))
        for column in df.columns
    }
    result = df.rename(columns=renamed).copy()
    for column in TEXT_COLUMNS:
        if column not in result.columns:
            result[column] = ""
        result[column] = result[column].map(clean_text)
    if "No." not in result.columns:
        result["No."] = range(1, len(result) + 1)
    result["No."] = pd.to_numeric(result["No."], errors="coerce")
    return result


def normalize_status(value: object) -> str:
    text = clean_text(value).lower()
    if not text:
        return "Not stated"

    uncertain = (
        "not confirmed",
        "no public",
        "not identified",
        "unconfirmed",
        "unknown",
        "pending",
        "awaiting",
    )
    negative = ("rejected", "withdrawn", "not approved", "unavailable")

    if any(term in text for term in negative):
        return "Not approved / unavailable"
    if any(term in text for term in uncertain):
        return "Unconfirmed / pending"
    if any(term in text for term in ("approved", "cleared", "clearance", "ce mark", "certified")):
        return "Approved / certified"
    if any(term in text for term in ("commercially available", "commercialized", "marketed", "on the market", "sales")):
        return "Commercially available"
    if any(term in text for term in ("development", "pre-commercial", "early stage", "early-stage", "prototype", "pilot")):
        return "Under development"
    if any(term in text for term in ("limited", "selected markets", "regional")):
        return "Limited availability"
    return "Other"


def detect_excel_header(source: object) -> int:
    preview = pd.read_excel(source, header=None, nrows=15)
    for index, row in preview.iterrows():
        values = {clean_text(value).lower() for value in row.tolist()}
        if "product name" in values and ("country of origin" in values or "product type" in values):
            return int(index)
    return 0


def read_data(source: object) -> pd.DataFrame:
    source_name = getattr(source, "name", str(source)).lower()
    if source_name.endswith(".csv"):
        raw = pd.read_csv(source)
    else:
        header_row = detect_excel_header(source)
        if hasattr(source, "seek"):
            source.seek(0)
        raw = pd.read_excel(source, header=header_row)

    data = normalize_columns(raw)
    data = data[data["Product Name"].str.len().gt(0)].copy()
    placeholder_mask = data["Product Name"].str.fullmatch(r"\[?\s*\.{2,}\s*\]?", na=False)
    data = data[~placeholder_mask].copy()

    data["US Category"] = data["U.S. Status"].map(normalize_status)
    data["Europe Category"] = data["Europe Status"].map(normalize_status)
    data["Middle East Category"] = data["Middle East Status"].map(normalize_status)
    data["Commercial Category"] = data["Commercial Status"].map(normalize_status)
    data["Primary Medical Field"] = data["Medical Field"].map(
        lambda value: first_nonempty(split_bullets(value))
    )
    return data.sort_values(["No.", "Product Name"], na_position="last").reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_manifest(path: str) -> pd.DataFrame:
    manifest_path = Path(path)
    if not manifest_path.is_file():
        return pd.DataFrame(columns=["No.", "Product Name", "Image Path", "Gallery Paths"])
    manifest = pd.read_csv(manifest_path)
    manifest = normalize_columns(manifest)
    return manifest[["No.", "Product Name", "Image Path", "Gallery Paths"]]


def attach_manifest(data: pd.DataFrame) -> pd.DataFrame:
    manifest = load_manifest(str(IMAGE_MANIFEST))
    if manifest.empty:
        return data

    result = data.copy()
    merge_key = "No." if result["No."].notna().any() else "Product Name"
    manifest_columns = [merge_key, "Image Path", "Gallery Paths"]
    result = result.merge(
        manifest[manifest_columns],
        on=merge_key,
        how="left",
        suffixes=("", "_manifest"),
    )

    for column in ("Image Path", "Gallery Paths"):
        manifest_column = f"{column}_manifest"
        if manifest_column not in result.columns:
            continue
        existing = result[column].fillna("").astype(str).str.strip()
        from_manifest = result[manifest_column].fillna("").astype(str).str.strip()
        result[column] = existing.where(existing.ne(""), from_manifest)
        result = result.drop(columns=[manifest_column])

    return result.sort_values(["No.", "Product Name"], na_position="last").reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_default_data(path: str) -> pd.DataFrame:
    return attach_manifest(read_data(Path(path)))


def validate_data(data: pd.DataFrame) -> list[str]:
    return [
        column
        for column in REQUIRED_COLUMNS
        if column not in data.columns or data[column].eq("").all()
    ]

# -----------------------------------------------------------------------------
# Images and display helpers
# -----------------------------------------------------------------------------
def candidate_image_paths(raw_path: str) -> list[Path]:
    cleaned = raw_path.strip().replace("\\", "/").lstrip("/")
    if not cleaned:
        return []
    filename = Path(cleaned).name
    candidates = [
        BASE_DIR / cleaned,
        IMAGE_DIR / filename,
        BASE_DIR / "images" / filename,
    ]
    unique: list[Path] = []
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved not in unique:
            unique.append(resolved)
    return unique


def resolve_gallery(row: pd.Series) -> list[Path]:
    raw = clean_text(row.get("Gallery Paths", ""))
    if not raw:
        raw = clean_text(row.get("Image Path", ""))

    resolved_images: list[Path] = []
    for part in (item.strip() for item in raw.split("|") if item.strip()):
        match = next((path for path in candidate_image_paths(part) if path.is_file()), None)
        if match and match not in resolved_images:
            resolved_images.append(match)

    # Last-resort match by product number, useful when a manifest path is stale.
    if not resolved_images and pd.notna(row.get("No.")) and IMAGE_DIR.is_dir():
        product_number = int(float(row["No."]))
        pattern = f"{product_number:02d}_*"
        resolved_images.extend(sorted(path.resolve() for path in IMAGE_DIR.glob(pattern) if path.is_file()))

    return resolved_images


def resolve_image(row: pd.Series) -> Path | None:
    gallery = resolve_gallery(row)
    return gallery[0] if gallery else None


def render_image_or_placeholder(row: pd.Series, height_hint: int = 230) -> None:
    image = resolve_image(row)
    if image:
        st.image(str(image), use_container_width=True)
        return

    st.markdown(
        f"""
        <div class="placeholder" style="min-height:{height_hint}px">
            <div><strong>Image coming soon</strong>Product #{escape(str(row.get('No.', '')))} has no linked image.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def website_url(value: str) -> str | None:
    candidate = first_nonempty(split_bullets(value), "")
    if not candidate or "no dedicated" in candidate.lower():
        return None
    if candidate.startswith(("https://", "http://")):
        return candidate
    return f"https://{candidate.strip()}"


def short_text(value: object, max_chars: int = 180) -> str:
    text = first_nonempty(split_bullets(value), "Not stated")
    return text if len(text) <= max_chars else text[: max_chars - 1].rstrip() + "…"


def options_from_bullets(series: pd.Series) -> list[str]:
    values: set[str] = set()
    for value in series:
        values.update(split_bullets(value))
    return sorted(values)


def contains_any(value: str, selected: list[str]) -> bool:
    return not selected or any(option.casefold() in value.casefold() for option in selected)


def badge_class(status: str) -> str:
    lowered = status.casefold()
    if "not approved" in lowered or "unavailable" in lowered:
        return "badge-negative"
    if "approved" in lowered or "commercially available" in lowered:
        return "badge-approved"
    if "certified" in lowered:
        return "badge-certified"
    if "development" in lowered or "limited" in lowered:
        return "badge-development"
    if "pending" in lowered or "unconfirmed" in lowered or "not stated" in lowered:
        return "badge-unconfirmed"
    return "badge-certified"


def badge_html(label: str, status: str) -> str:
    return (
        f'<span class="badge {badge_class(status)}">'
        f"{escape(label)} · {escape(status)}</span>"
    )


def info_panel(title: str, value: str) -> None:
    st.markdown(
        f'<div class="info-panel"><h4>{escape(title)}</h4><p>{escape(value or "Not stated")}</p></div>',
        unsafe_allow_html=True,
    )


def render_product_card(row: pd.Series, key_prefix: str) -> None:
    with st.container(border=True):
        render_image_or_placeholder(row)
        st.markdown(
            f'<div class="product-title">{escape(row["Product Name"])}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<div class="product-meta">{escape(row["Country of Origin"])} · '
            f'{escape(first_nonempty(split_bullets(row["Product Type"])))}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            badge_html("Commercial", row["Commercial Category"])
            + badge_html("U.S.", row["US Category"]),
            unsafe_allow_html=True,
        )
        st.markdown('<div class="section-label">Innovation</div>', unsafe_allow_html=True)
        st.write(short_text(row["Innovation"], 170))
        if st.button(
            "View full profile",
            key=f"{key_prefix}_{row['No.']}_{row['Product Name']}",
            use_container_width=True,
            type="primary",
        ):
            st.session_state.selected_product_no = row["No."]
            st.rerun()


def render_profile(row: pd.Series) -> None:
    gallery = resolve_gallery(row)
    left, right = st.columns([1.05, 1.45], gap="large")

    with left:
        if gallery:
            product_no = int(float(row["No."])) if pd.notna(row["No."]) else 0
            state_key = f"gallery_index_{product_no}"
            st.session_state.setdefault(state_key, 0)
            current = st.session_state[state_key] % len(gallery)
            st.image(str(gallery[current]), use_container_width=True)
            if len(gallery) > 1:
                previous, count, next_button = st.columns([1, 2, 1])
                if previous.button("←", key=f"previous_{state_key}", use_container_width=True):
                    st.session_state[state_key] = (current - 1) % len(gallery)
                    st.rerun()
                count.markdown(
                    f"<div style='text-align:center;color:#667085;padding:.55rem'>{current + 1} of {len(gallery)}</div>",
                    unsafe_allow_html=True,
                )
                if next_button.button("→", key=f"next_{state_key}", use_container_width=True):
                    st.session_state[state_key] = (current + 1) % len(gallery)
                    st.rerun()
        else:
            render_image_or_placeholder(row, height_hint=360)

    with right:
        st.markdown('<div class="eyebrow">Product intelligence profile</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="profile-title">{escape(row["Product Name"])}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<div class="product-meta">{escape(row["Country of Origin"])} · '
            f'{escape(first_nonempty(split_bullets(row["Product Type"])))}</div>',
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
        st.write(short_text(row["Innovation"], 520))

        meta_left, meta_right = st.columns(2)
        with meta_left:
            st.markdown("**Manufacturer / developer**")
            st.write(row["Manufacturer/Developer"] or "Not stated")
        with meta_right:
            st.markdown("**Medical field**")
            st.write(row["Medical Field"] or "Not stated")

        url = website_url(row["Website"])
        if url:
            st.link_button("Open product or manufacturer website ↗", url, use_container_width=True)

    st.markdown("### Research profile")
    overview, clinical, market, regulatory = st.tabs(
        ["Overview", "Clinical value", "Market", "Regulatory"]
    )

    with overview:
        col1, col2 = st.columns(2)
        with col1:
            info_panel("Innovation", row["Innovation"])
            st.write("")
            info_panel("Materials", row["Material Used"])
        with col2:
            info_panel("Product type", row["Product Type"])
            st.write("")
            info_panel("Medical field", row["Medical Field"])

    with clinical:
        col1, col2, col3 = st.columns(3)
        with col1:
            info_panel("Hospital need", row["Hospital Need"])
        with col2:
            info_panel("Advantages", row["Advantages"])
        with col3:
            info_panel("Limitations", row["Disadvantages/Limitations"])

    with market:
        col1, col2, col3 = st.columns(3)
        with col1:
            info_panel("Commercial status", row["Commercial Status"])
        with col2:
            info_panel("Market potential", row["Market Potential"])
        with col3:
            info_panel("Recommendation", row["Recommendation"])

    with regulatory:
        for label, value in (
            ("United States", row["U.S. Status"]),
            ("Europe", row["Europe Status"]),
            ("Middle East", row["Middle East Status"]),
        ):
            st.markdown(f"**{label}**")
            st.write(value or "Not stated")
            st.divider()

# -----------------------------------------------------------------------------
# Data source and global filters
# -----------------------------------------------------------------------------
st.session_state.setdefault("selected_product_no", None)

with st.sidebar:
    st.markdown("## 🩺 Device Explorer")
    st.caption("Medical technology intelligence")
    page = st.radio(
        "Navigation",
        ["Dashboard", "Explore", "Map", "Compare", "Insights", "Data"],
        label_visibility="collapsed",
    )

    st.divider()
    st.markdown("### Data source")
    uploaded_file = st.file_uploader(
        "Upload Excel or CSV",
        type=["xlsx", "xls", "csv"],
        help="The uploaded file is used only for this browser session.",
    )

    try:
        data = attach_manifest(read_data(uploaded_file)) if uploaded_file else load_default_data(str(DEFAULT_DATA))
    except Exception as exc:
        st.error(f"Could not read the data file: {exc}")
        st.stop()

    missing = validate_data(data)
    if missing:
        st.error("Missing required columns: " + ", ".join(missing))
        st.stop()

    image_count = sum(resolve_image(row) is not None for _, row in data.iterrows())
    st.success(f"{len(data)} products · {image_count} with images")

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
    searchable_columns = [column for column in TEXT_COLUMNS if column in filtered.columns]
    search_mask = filtered[searchable_columns].apply(
        lambda column: column.str.contains(search, case=False, na=False)
    ).any(axis=1)
    filtered = filtered[search_mask]
if countries:
    filtered = filtered[filtered["Country of Origin"].isin(countries)]
if product_types:
    filtered = filtered[filtered["Product Type"].map(lambda value: contains_any(value, product_types))]
if fields:
    filtered = filtered[filtered["Medical Field"].map(lambda value: contains_any(value, fields))]
if commercial:
    filtered = filtered[filtered["Commercial Category"].isin(commercial)]
if us_status:
    filtered = filtered[filtered["US Category"].isin(us_status)]
if eu_status:
    filtered = filtered[filtered["Europe Category"].isin(eu_status)]

filtered = filtered.sort_values(["No.", "Product Name"], na_position="last").reset_index(drop=True)

# -----------------------------------------------------------------------------
# Dedicated product profile
# -----------------------------------------------------------------------------
if st.session_state.selected_product_no is not None:
    selected_number = pd.to_numeric(st.session_state.selected_product_no, errors="coerce")
    selected_rows = data[data["No."] == selected_number]
    if selected_rows.empty:
        st.session_state.selected_product_no = None
        st.rerun()

    if st.button("← Back to product explorer"):
        st.session_state.selected_product_no = None
        st.rerun()

    with st.container(border=True):
        render_profile(selected_rows.iloc[0])

    st.divider()
    st.caption(
        "Research support tool only. Verify regulatory and commercial information before procurement or clinical decisions."
    )
    st.stop()

# -----------------------------------------------------------------------------
# Main views
# -----------------------------------------------------------------------------
if page == "Dashboard":
    st.markdown(
        """
        <div class="hero">
            <div class="eyebrow">Emerging medical technology intelligence</div>
            <h1>Discover, compare, and evaluate promising medical devices.</h1>
            <p>Explore a curated research portfolio through product profiles, regulatory signals, market-readiness indicators, and geographic insights.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric1, metric2, metric3, metric4 = st.columns(4)
    metric1.metric("Products", len(filtered))
    metric2.metric("Countries represented", filtered["Country of Origin"].nunique())
    metric3.metric("Medical fields", filtered["Primary Medical Field"].nunique())
    metric4.metric(
        "Commercially available",
        int((filtered["Commercial Category"] == "Commercially available").sum()),
    )

    st.markdown("### Featured products")
    if filtered.empty:
        st.info("No products match the current filters.")
    else:
        featured = filtered.head(3)
        columns = st.columns(3, gap="large")
        for column, (_, row) in zip(columns, featured.iterrows()):
            with column:
                render_product_card(row, "dashboard")

    st.markdown("### Portfolio snapshot")
    chart_col, summary_col = st.columns([1.45, 1], gap="large")
    with chart_col:
        if not filtered.empty:
            map_data = filtered.groupby("Country of Origin", as_index=False).agg(
                Products=("Product Name", "count"),
                Devices=("Product Name", lambda values: "<br>".join(values)),
            )
            figure = px.scatter_geo(
                map_data,
                locations="Country of Origin",
                locationmode="country names",
                size="Products",
                hover_name="Country of Origin",
                hover_data={"Products": True, "Devices": True},
                projection="natural earth",
            )
            figure.update_geos(
                showframe=False,
                showcoastlines=True,
                coastlinecolor="#cbd5e1",
                landcolor="#f8fafc",
                bgcolor="rgba(0,0,0,0)",
            )
            figure.update_layout(
                height=410,
                margin=dict(l=0, r=0, t=0, b=0),
                paper_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(figure, use_container_width=True)
    with summary_col:
        with st.container(border=True):
            st.markdown("#### Current selection")
            st.write(f"**{len(filtered)}** products match the active filters.")
            st.write(f"**{image_count}** products in the full dataset currently have linked imagery.")
            st.write(
                "Use **Explore** for product cards, **Map** for geography, and **Compare** for side-by-side review."
            )

elif page == "Explore":
    heading, controls = st.columns([2.4, 1])
    with heading:
        st.markdown("## Product explorer")
        st.caption(f"Showing {len(filtered)} of {len(data)} products")
    with controls:
        sort_choice = st.selectbox(
            "Sort products",
            ["Product number", "Product name", "Country", "Commercial status"],
        )

    sort_map = {
        "Product number": ["No.", "Product Name"],
        "Product name": ["Product Name"],
        "Country": ["Country of Origin", "Product Name"],
        "Commercial status": ["Commercial Category", "Product Name"],
    }
    sorted_products = filtered.sort_values(sort_map[sort_choice], na_position="last")

    if sorted_products.empty:
        st.info("No products match the current filters.")
    else:
        product_rows = list(sorted_products.iterrows())
        for start in range(0, len(product_rows), 3):
            columns = st.columns(3, gap="large")
            for column, (_, row) in zip(columns, product_rows[start : start + 3]):
                with column:
                    render_product_card(row, "explore")

elif page == "Map":
    st.markdown("## Global device map")
    st.caption("Bubble size represents the number of products associated with each country.")
    if filtered.empty:
        st.info("No products match the current filters.")
    else:
        map_data = filtered.groupby("Country of Origin", as_index=False).agg(
            Products=("Product Name", "count"),
            Devices=("Product Name", lambda values: "<br>".join(values)),
        )
        figure = px.scatter_geo(
            map_data,
            locations="Country of Origin",
            locationmode="country names",
            size="Products",
            hover_name="Country of Origin",
            hover_data={"Products": True, "Devices": True},
            projection="natural earth",
        )
        figure.update_geos(
            showframe=False,
            showcoastlines=True,
            coastlinecolor="#cbd5e1",
            landcolor="#f8fafc",
            bgcolor="rgba(0,0,0,0)",
        )
        figure.update_layout(
            height=650,
            margin=dict(l=0, r=0, t=10, b=0),
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(figure, use_container_width=True)

elif page == "Compare":
    st.markdown("## Compare products")
    selected_products = st.multiselect(
        "Choose up to four products",
        filtered["Product Name"].tolist(),
        max_selections=4,
    )

    if not selected_products:
        st.info("Select two to four products to create a side-by-side comparison.")
    else:
        selected_data = filtered[filtered["Product Name"].isin(selected_products)]
        columns = st.columns(len(selected_data), gap="large")
        for column, (_, row) in zip(columns, selected_data.iterrows()):
            with column:
                with st.container(border=True):
                    render_image_or_placeholder(row, height_hint=170)
                    st.markdown(f"**{row['Product Name']}**")
                    st.caption(f"{row['Country of Origin']} · {row['Primary Medical Field']}")

        comparison_columns = [
            "Product Name",
            "Product Type",
            "Country of Origin",
            "Medical Field",
            "Innovation",
            "Material Used",
            "Advantages",
            "Disadvantages/Limitations",
            "Commercial Category",
            "US Category",
            "Europe Category",
            "Middle East Category",
            "Recommendation",
        ]
        comparison = selected_data[comparison_columns].set_index("Product Name").T
        st.dataframe(comparison, use_container_width=True)

elif page == "Insights":
    st.markdown("## Portfolio insights")
    if filtered.empty:
        st.info("No products match the current filters.")
    else:
        chart1, chart2 = st.columns(2, gap="large")
        with chart1:
            country_counts = (
                filtered["Country of Origin"]
                .value_counts()
                .rename_axis("Country")
                .reset_index(name="Products")
            )
            country_chart = px.bar(
                country_counts,
                x="Country",
                y="Products",
                title="Products by country",
            )
            country_chart.update_layout(margin=dict(l=10, r=10, t=55, b=10))
            st.plotly_chart(country_chart, use_container_width=True)

        with chart2:
            status_counts = (
                filtered["Commercial Category"]
                .value_counts()
                .rename_axis("Status")
                .reset_index(name="Products")
            )
            status_chart = px.pie(
                status_counts,
                names="Status",
                values="Products",
                hole=.52,
                title="Commercial readiness",
            )
            st.plotly_chart(status_chart, use_container_width=True)

        field_counts = (
            filtered["Primary Medical Field"]
            .value_counts()
            .head(12)
            .rename_axis("Medical field")
            .reset_index(name="Products")
        )
        field_chart = px.bar(
            field_counts,
            x="Products",
            y="Medical field",
            orientation="h",
            title="Leading medical fields",
        )
        st.plotly_chart(field_chart, use_container_width=True)

else:
    st.markdown("## Research data")
    display_columns = [
        column
        for column in [
            "No.",
            "Product Name",
            "Product Type",
            "Country of Origin",
            "Medical Field",
            "Commercial Category",
            "US Category",
            "Europe Category",
            "Middle East Category",
            "Website",
        ]
        if column in filtered.columns
    ]
    st.dataframe(filtered[display_columns], use_container_width=True, hide_index=True)
    csv_data = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered data as CSV",
        data=csv_data,
        file_name="filtered_medical_devices.csv",
        mime="text/csv",
    )

st.divider()
st.caption(
    "Research support tool only. Regulatory and commercial information should be independently verified before procurement or clinical decisions."
)
