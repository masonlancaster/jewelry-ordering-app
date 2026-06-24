from __future__ import annotations

import math
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import streamlit as st

from add_to_cart_dluxca import (
    DEFAULT_SITE,
    CartUrlError,
    SkippedItem,
    build_shopify_cart_url,
    find_column,
    is_yes,
    load_dluxca_items_from_excel,
    normalize_site,
    parse_number,
    read_xlsx_rows,
    site_host,
)


st.set_page_config(page_title="DLUXCA Order Helper", layout="wide")


st.markdown(
    """
    <style>
    :root {
        --ink: #1f2933;
        --muted: #667085;
        --line: #d9e2ec;
        --panel: #ffffff;
        --soft: #f6f8fb;
        --accent: #0f766e;
        --accent-dark: #115e59;
        --warn: #b45309;
    }

    .stApp {
        background:
            linear-gradient(180deg, #f8fafc 0%, #eef3f8 100%);
        color: var(--ink);
    }

    .block-container {
        max-width: 1180px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    [data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid var(--line);
    }

    .app-header {
        background: #ffffff;
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 1.5rem 1.6rem;
        margin-bottom: 1rem;
        box-shadow: 0 12px 34px rgba(31, 41, 51, 0.07);
    }

    .app-kicker {
        color: var(--accent-dark);
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0;
        text-transform: uppercase;
        margin-bottom: 0.4rem;
    }

    .app-title {
        color: var(--ink);
        font-size: 2.05rem;
        line-height: 1.15;
        font-weight: 750;
        letter-spacing: 0;
        margin: 0;
    }

    .app-subtitle {
        color: var(--muted);
        font-size: 1rem;
        line-height: 1.55;
        max-width: 760px;
        margin-top: 0.65rem;
    }

    .section-panel {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 1.15rem;
        box-shadow: 0 8px 24px rgba(31, 41, 51, 0.05);
        margin-bottom: 1rem;
    }

    .panel-title {
        color: var(--ink);
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 0.25rem;
    }

    .panel-copy {
        color: var(--muted);
        font-size: 0.92rem;
        line-height: 1.45;
        margin-bottom: 0.85rem;
    }

    .status-box {
        border-left: 4px solid var(--accent);
        background: #eefaf7;
        border-radius: 6px;
        padding: 0.9rem 1rem;
        color: #164e45;
        margin: 0.8rem 0 1rem;
    }

    .warning-box {
        border-left: 4px solid var(--warn);
        background: #fff7ed;
        border-radius: 6px;
        padding: 0.9rem 1rem;
        color: #7c2d12;
        margin: 0.8rem 0 1rem;
    }

    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 0.9rem 1rem;
        box-shadow: 0 8px 20px rgba(31, 41, 51, 0.05);
    }

    div[data-testid="stMetricLabel"] p {
        color: var(--muted);
        font-size: 0.82rem;
    }

    div[data-testid="stMetricValue"] {
        color: var(--ink);
    }

    .stButton > button,
    .stLinkButton > a {
        border-radius: 6px;
        font-weight: 700;
        border: 1px solid var(--accent);
    }

    .stButton > button[kind="primary"],
    .stLinkButton > a {
        background: var(--accent);
        color: #ffffff;
    }

    .stButton > button[kind="primary"]:hover,
    .stLinkButton > a:hover {
        background: var(--accent-dark);
        border-color: var(--accent-dark);
        color: #ffffff;
    }

    textarea {
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace;
        font-size: 0.86rem !important;
    }

    h2, h3 {
        letter-spacing: 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


class AppArgs:
    link_column = "Link"
    quantity_column = "Quantity Desired"
    on_floor_column = "On Floor (#)"
    back_stock_column = "Back Stock (#)"
    reorder_column = "Reorder Needed"
    sku_column = "Sku"
    name_column = "Name"


@dataclass(frozen=True)
class ManualOrder:
    row_number: int
    store: str
    name: str
    sku: str
    link: str
    quantity: int


def skipped_table_rows(skipped: list[SkippedItem]) -> list[dict[str, object]]:
    return [
        {
            "SKU": item.sku,
            "Name": item.name,
            "Quantity Needed": item.quantity,
            "URL": item.link,
            "Reason": item.reason,
        }
        for item in skipped
    ]


def manual_order_table_rows(orders: list[ManualOrder]) -> list[dict[str, object]]:
    return [
        {
            "SKU": order.sku,
            "Name": order.name,
            "Quantity Needed": order.quantity,
            "URL": order.link,
        }
        for order in orders
    ]


def collect_other_store_orders(path: Path, args: AppArgs) -> dict[str, list[ManualOrder]]:
    rows = read_xlsx_rows(path)
    required_columns = [
        args.link_column,
        args.quantity_column,
        args.on_floor_column,
        args.back_stock_column,
        args.reorder_column,
    ]
    missing_columns: set[str] = set()
    grouped: dict[str, list[ManualOrder]] = defaultdict(list)

    for row_number, row in rows:
        keys = {column: find_column(row, column) for column in required_columns}
        for column, key in keys.items():
            if key is None:
                missing_columns.add(column)
        if missing_columns:
            continue

        if not is_yes(row.get(keys[args.reorder_column] or "", "")):
            continue

        link = row.get(keys[args.link_column] or "", "").strip()
        store = site_host(link)
        if not store or store == "dluxca.com":
            continue

        quantity_desired = parse_number(
            row.get(keys[args.quantity_column] or "", ""), args.quantity_column, row_number
        )
        on_floor = parse_number(
            row.get(keys[args.on_floor_column] or "", ""), args.on_floor_column, row_number
        )
        back_stock = parse_number(
            row.get(keys[args.back_stock_column] or "", ""), args.back_stock_column, row_number
        )
        quantity = int(math.ceil(quantity_desired - on_floor - back_stock))
        if quantity <= 0:
            continue

        sku_key = find_column(row, args.sku_column)
        name_key = find_column(row, args.name_column)
        grouped[store].append(
            ManualOrder(
                row_number=row_number,
                store=store,
                name=row.get(name_key, "").strip() if name_key else "",
                sku=row.get(sku_key, "").strip() if sku_key else "",
                link=link,
                quantity=quantity,
            )
        )

    if missing_columns:
        raise CartUrlError(f"Missing required Excel column(s): {', '.join(sorted(missing_columns))}")
    return dict(sorted(grouped.items()))


def process_upload(uploaded_file, site: str):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        temp_path = Path(temp_file.name)

    try:
        args = AppArgs()
        items = load_dluxca_items_from_excel(temp_path, args)
        other_store_orders = collect_other_store_orders(temp_path, args)
        result = build_shopify_cart_url(normalize_site(site), items)
        return items, result, other_store_orders
    finally:
        temp_path.unlink(missing_ok=True)


if "last_result" not in st.session_state:
    st.session_state.last_result = None
if "last_items" not in st.session_state:
    st.session_state.last_items = None
if "last_file_name" not in st.session_state:
    st.session_state.last_file_name = None
if "last_other_store_orders" not in st.session_state:
    st.session_state.last_other_store_orders = None


st.markdown(
    """
    <div class="app-header">
        <div class="app-kicker">Inventory ordering</div>
        <h1 class="app-title">DLUXCA Order Helper</h1>
        <div class="app-subtitle">
            Upload one monday.com inventory export, generate the DLUXCA cart URL,
            and review any items that need manual attention before ordering.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Settings")
    site = st.text_input("DLUXCA website", DEFAULT_SITE)
    st.caption("The app uses Shopify product data to build a cart URL. It does not log in or place the order.")

left_col, right_col = st.columns([1.05, 0.95], gap="large")

with left_col:
    st.markdown(
        """
        <div class="section-panel">
            <div class="panel-title">Upload Inventory Export</div>
            <div class="panel-copy">
                Choose the monday.com Excel export. The app will only use DLUXCA
                rows marked for reorder.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader("Inventory .xlsx file", type=["xlsx"], label_visibility="collapsed")
    generate_clicked = st.button(
        "Generate cart URL",
        type="primary",
        use_container_width=True,
        disabled=uploaded_file is None,
    )

with right_col:
    st.markdown(
        """
        <div class="section-panel">
            <div class="panel-title">What The App Checks</div>
            <div class="panel-copy">
                It includes rows where Reorder Needed is yes, calculates the
                needed quantity, builds the DLUXCA cart URL, and groups the
                remaining supplier orders for manual review.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

if uploaded_file is None:
    st.markdown(
        """
        <div class="status-box">
            Upload an inventory export to generate the cart URL.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

if st.session_state.last_file_name != uploaded_file.name and not generate_clicked:
    st.session_state.last_items = None
    st.session_state.last_result = None
    st.session_state.last_other_store_orders = None

if generate_clicked:
    with st.spinner("Reading workbook and resolving DLUXCA products..."):
        try:
            items, result, other_store_orders = process_upload(uploaded_file, site)
            st.session_state.last_items = items
            st.session_state.last_result = result
            st.session_state.last_other_store_orders = other_store_orders
            st.session_state.last_file_name = uploaded_file.name
        except CartUrlError as exc:
            st.error(str(exc))
            st.stop()
        except Exception as exc:
            st.error(f"Something went wrong while processing the file: {exc}")
            st.stop()

items = st.session_state.last_items
result = st.session_state.last_result
other_store_orders = st.session_state.last_other_store_orders or {}

if result is None or items is None:
    st.markdown(
        """
        <div class="status-box">
            File loaded. Click Generate cart URL when ready.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

st.markdown(
    f"""
    <div class="status-box">
        Cart URL generated from <strong>{st.session_state.last_file_name}</strong>.
    </div>
    """,
    unsafe_allow_html=True,
)

metric_cols = st.columns(4)
metric_cols[0].metric("DLUXCA rows found", len(items))
metric_cols[1].metric("Rows added", result.resolved_count)
metric_cols[2].metric("Total quantity", result.resolved_quantity)
metric_cols[3].metric(
    "Manual review",
    len(result.skipped) + sum(len(orders) for orders in other_store_orders.values()),
)

st.subheader("Cart URL")
st.text_area("Copy this URL", result.cart_url, height=120, label_visibility="collapsed")
st.link_button("Open DLUXCA Cart", result.cart_url, use_container_width=False)

st.subheader("Manual Review Orders")
if result.skipped:
    st.markdown("#### DLUXCA skipped")
    st.markdown(
        """
        <div class="warning-box">
            These DLUXCA rows were not added to the cart URL and need manual review.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.dataframe(skipped_table_rows(result.skipped), use_container_width=True, hide_index=True)
else:
    st.markdown("#### DLUXCA skipped")
    st.markdown(
        """
        <div class="status-box">
            No skipped DLUXCA orders.
        </div>
        """,
        unsafe_allow_html=True,
    )

if other_store_orders:
    st.markdown("#### Other stores")
    for store, orders in other_store_orders.items():
        st.markdown(f"##### {store}")
        st.dataframe(manual_order_table_rows(orders), use_container_width=True, hide_index=True)
else:
    st.markdown("#### Other stores")
    st.markdown(
        """
        <div class="status-box">
            No reorder rows found for other stores.
        </div>
        """,
        unsafe_allow_html=True,
    )
