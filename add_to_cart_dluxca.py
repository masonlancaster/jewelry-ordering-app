from __future__ import annotations

import argparse
import json
import math
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from urllib.error import URLError
from urllib.parse import parse_qs, urlparse
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET
from zipfile import ZipFile


DEFAULT_SITE = "https://dluxca.com"
DEFAULT_INPUT_DIR = "monday_files"
DEFAULT_OUTPUT_DIR = "new_files"

COLOR_PHRASES = (
    "dark purple",
    "dark blue",
    "light blue",
    "navy blue",
    "dark red",
    "olive green",
    "white gold",
    "rose gold",
    "yellow gold",
    "clear",
    "diamond",
    "emerald",
    "alexandrite",
    "ruby",
    "garnet",
    "peridot",
    "citrine",
    "fuchsia",
    "purple",
    "yellow",
    "silver",
    "gold",
    "green",
    "blue",
    "pink",
    "red",
    "black",
    "white",
    "peach",
)


@dataclass(frozen=True)
class BatchItem:
    row_number: int
    name: str
    sku: str
    link: str
    quantity: int


@dataclass(frozen=True)
class SkippedItem:
    row_number: int
    name: str
    sku: str
    link: str
    quantity: int
    reason: str


@dataclass(frozen=True)
class CartUrlResult:
    cart_url: str
    resolved_count: int
    resolved_quantity: int
    skipped: list[SkippedItem]


class CartUrlError(RuntimeError):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Read monday.com inventory .xlsx files and write DLUXCA Shopify cart "
            "permalink text files."
        )
    )
    parser.add_argument("--site", default=DEFAULT_SITE, help="DLUXCA storefront base URL.")
    parser.add_argument(
        "--excel",
        help="Optional path to one .xlsx file. By default, every .xlsx in monday_files/ is processed.",
    )
    parser.add_argument(
        "--input-dir",
        default=DEFAULT_INPUT_DIR,
        help='Directory to read .xlsx files from when --excel is not supplied. Default: "monday_files".',
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help='Directory to write cart URL and skipped files to. Default: "new_files".',
    )
    parser.add_argument("--link-column", default="Link")
    parser.add_argument("--quantity-column", default="Quantity Desired")
    parser.add_argument("--on-floor-column", default="On Floor (#)")
    parser.add_argument("--back-stock-column", default="Back Stock (#)")
    parser.add_argument("--reorder-column", default="Reorder Needed")
    parser.add_argument("--sku-column", default="Sku")
    parser.add_argument("--name-column", default="Name")
    parser.add_argument(
        "--limit",
        type=int,
        help="Limit the number of matching DLUXCA rows processed. Useful for testing.",
    )
    return parser.parse_args()


def normalize_site(site: str) -> str:
    site = site.strip()
    if not site.startswith(("http://", "https://")):
        site = f"https://{site}"
    return site.rstrip("/")


def normalized_sku(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def normalized_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def site_host(site: str) -> str:
    return urlparse(site.strip()).netloc.lower().removeprefix("www.")


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for char in letters:
        index = index * 26 + ord(char.upper()) - 64
    return index - 1


def read_shared_strings(zip_file: ZipFile) -> list[str]:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    return [
        "".join(text.text or "" for text in si.findall(".//m:t", ns))
        for si in root.findall("m:si", ns)
    ]


def cell_value(cell: ET.Element, strings: list[str]) -> str:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    if cell.get("t") == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//m:t", ns)).strip()
    value = cell.find("m:v", ns)
    if value is None:
        return ""
    raw = (value.text or "").strip()
    if cell.get("t") == "s" and raw.isdigit() and int(raw) < len(strings):
        return strings[int(raw)].strip()
    return raw


def read_xlsx_rows(path: Path) -> list[tuple[int, dict[str, str]]]:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(path) as zip_file:
        strings = read_shared_strings(zip_file)
        root = ET.fromstring(zip_file.read("xl/worksheets/sheet1.xml"))
        raw_rows: list[tuple[int, dict[int, str]]] = []
        for row in root.findall("m:sheetData/m:row", ns):
            row_number = int(row.get("r") or len(raw_rows) + 1)
            values = {
                column_index(cell.get("r") or ""): cell_value(cell, strings)
                for cell in row.findall("m:c", ns)
            }
            raw_rows.append((row_number, values))

    header_row_index = None
    headers: dict[int, str] = {}
    for index, (_row_number, values) in enumerate(raw_rows):
        normalized_values = {normalized_header(value): col for col, value in values.items()}
        if "link" in normalized_values:
            header_row_index = index
            headers = values
            break
    if header_row_index is None:
        raise CartUrlError("Could not find a header row containing a Link column.")

    rows: list[tuple[int, dict[str, str]]] = []
    for row_number, values in raw_rows[header_row_index + 1 :]:
        mapped = {header: values.get(col, "") for col, header in headers.items() if header}
        if any(mapped.values()):
            rows.append((row_number, mapped))
    return rows


def read_xlsx_cell(path: Path, cell_ref: str) -> str:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(path) as zip_file:
        strings = read_shared_strings(zip_file)
        root = ET.fromstring(zip_file.read("xl/worksheets/sheet1.xml"))
        cell = root.find(f'.//m:c[@r="{cell_ref}"]', ns)
        if cell is None:
            return ""
        return cell_value(cell, strings)


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", value).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "dluxca_order"


def output_base_for_excel(path: Path) -> str:
    workbook_name = read_xlsx_cell(path, "A1")
    return safe_filename(workbook_name or path.stem)


def parse_number(value: str, column: str, row_number: int) -> float:
    if value is None or str(value).strip() == "":
        return 0
    try:
        number = float(str(value).strip())
    except ValueError as exc:
        raise CartUrlError(f"Invalid number {value!r} in {column!r} on row {row_number}.") from exc
    if not math.isfinite(number):
        raise CartUrlError(f"Invalid number {value!r} in {column!r} on row {row_number}.")
    return number


def is_yes(value: str) -> bool:
    return str(value or "").strip().lower() in {"yes", "y", "true", "1"}


def find_column(row: dict[str, str], wanted: str) -> str | None:
    target = normalized_header(wanted)
    for key in row:
        if normalized_header(key) == target:
            return key
    return None


def load_dluxca_items_from_excel(path: Path, args: argparse.Namespace) -> list[BatchItem]:
    rows = read_xlsx_rows(path)
    items: list[BatchItem] = []
    required_columns = [
        args.link_column,
        args.quantity_column,
        args.on_floor_column,
        args.back_stock_column,
        args.reorder_column,
    ]
    missing_columns: set[str] = set()

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
        if site_host(link) != "dluxca.com":
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
        items.append(
            BatchItem(
                row_number=row_number,
                name=row.get(name_key, "").strip() if name_key else "",
                sku=row.get(sku_key, "").strip() if sku_key else "",
                link=link,
                quantity=quantity,
            )
        )

    if missing_columns:
        raise CartUrlError(f"Missing required Excel column(s): {', '.join(sorted(missing_columns))}")
    return items


def product_json_url_from_link(site: str, link: str) -> str:
    parsed = urlparse(link)
    match = re.search(r"/products/([^/?#]+)", parsed.path)
    if not match:
        raise CartUrlError(f"DLUXCA link is not a Shopify product URL: {link}")
    return f"{site}/products/{match.group(1)}.js"


def requested_variant_id_from_link(link: str) -> str | None:
    values = parse_qs(urlparse(link).query).get("variant")
    return values[0] if values else None


def clean_variant_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def name_tokens(value: str) -> list[str]:
    name_hint = str(value or "").split(":", 1)[-1]
    return re.findall(r"[a-z]+|[0-9]+", clean_variant_text(name_hint))


def extract_color_requests(name: str) -> list[str]:
    text = f" {str(name or '').split(':', 1)[-1].lower()} "
    found: list[tuple[int, str]] = []
    consumed: list[tuple[int, int]] = []
    for phrase in COLOR_PHRASES:
        pattern = r"(?<![a-z])" + re.escape(phrase) + r"(?![a-z])"
        for match in re.finditer(pattern, text):
            span = match.span()
            if any(not (span[1] <= taken[0] or span[0] >= taken[1]) for taken in consumed):
                continue
            found.append((span[0], phrase))
            consumed.append(span)
    colors: list[str] = []
    for _position, color in sorted(found):
        if color not in colors:
            colors.append(color)
    return colors


def split_quantity(quantity: int, parts: int) -> list[int]:
    base = quantity // parts
    remainder = quantity % parts
    return [base + (1 if index < remainder else 0) for index in range(parts)]


def variant_text(variant: dict) -> str:
    return clean_variant_text(
        " ".join(
            str(variant.get(key) or "")
            for key in ("title", "option1", "option2", "option3", "public_title", "sku")
        )
    )


def variant_matches_color(variant: dict, color: str) -> bool:
    text = variant_text(variant)
    compact_color = clean_variant_text(color)
    if compact_color in text:
        return True
    if color == "diamond":
        return "clear" in text
    return False


def fetch_json_url(url: str) -> dict:
    request = Request(
        url,
        headers={
            "Accept": "application/json",
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
            ),
        },
    )
    try:
        with urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except URLError as exc:
        raise CartUrlError(f"Could not fetch Shopify product data from {url}: {exc}") from exc


def variant_matches_name(variant: dict, item: BatchItem) -> bool:
    tokens = name_tokens(item.name)
    if not tokens:
        return False
    title = clean_variant_text(
        " ".join(
            str(variant.get(key) or "")
            for key in ("title", "option1", "option2", "option3", "public_title")
        )
    )
    for token in tokens:
        useful_token = len(token) >= 1 if token.isdigit() else len(token) >= 4
        if useful_token and token in title:
            return True
    return False


def select_shopify_variant(product: dict, item: BatchItem) -> dict:
    variants = product.get("variants") or []
    requested_variant_id = requested_variant_id_from_link(item.link)
    if requested_variant_id:
        for variant in variants:
            if str(variant.get("id")) == requested_variant_id:
                return variant

    item_sku = normalized_sku(item.sku)
    for variant in variants:
        if normalized_sku(str(variant.get("sku") or "")) == item_sku:
            return variant

    name_matches = [variant for variant in variants if variant_matches_name(variant, item)]
    if len(name_matches) == 1:
        return name_matches[0]
    if len(name_matches) > 1:
        raise CartUrlError(
            f"Row {item.row_number} is ambiguous: {len(name_matches)} variants match "
            f"{item.name or item.sku!r}."
        )

    if len(variants) > 1:
        raise CartUrlError(
            f"Row {item.row_number} is ambiguous: product has {len(variants)} variants "
            "but the URL does not specify one."
        )

    for variant in variants:
        if variant.get("available", True):
            return variant

    if variants:
        return variants[0]
    raise CartUrlError(f"Could not find any variants for row {item.row_number}: {item.link}")


def select_color_allocations(product: dict, item: BatchItem, colors: list[str]) -> list[tuple[dict, int, str]]:
    variants = product.get("variants") or []
    allocations: list[tuple[dict, int, str]] = []
    quantities = split_quantity(item.quantity, len(colors))
    for color, quantity in zip(colors, quantities):
        matches = [variant for variant in variants if variant_matches_color(variant, color)]
        available_matches = [variant for variant in matches if variant.get("available", True)]
        if len(available_matches) != 1:
            raise CartUrlError(
                f"Row {item.row_number} color {color!r} matched "
                f"{len(available_matches)} available variants."
            )
        allocations.append((available_matches[0], quantity, f"color {color}"))
    return allocations


def select_shopify_variant_allocations(product: dict, item: BatchItem) -> list[tuple[dict, int, str]]:
    requested_variant_id = requested_variant_id_from_link(item.link)
    colors = extract_color_requests(item.name)
    if requested_variant_id:
        variant = select_shopify_variant(product, item)
        if colors and len(product.get("variants") or []) > 1:
            try:
                color_allocations = select_color_allocations(product, item, colors)
                color_variant_ids = {str(variant.get("id")) for variant, _quantity, _reason in color_allocations}
                if str(variant.get("id")) not in color_variant_ids:
                    return color_allocations
            except CartUrlError:
                pass
        return [(variant, item.quantity, "explicit URL variant")]

    variants = product.get("variants") or []
    if colors and len(variants) > 1:
        return select_color_allocations(product, item, colors)

    variant = select_shopify_variant(product, item)
    return [(variant, item.quantity, "single variant")]


def format_skipped_items(skipped: list[SkippedItem]) -> str:
    return "\n".join(
        f"SKU: {item.sku}\n"
        f"Name: {item.name}\n"
        f"Quantity Needed: {item.quantity}\n"
        f"URL: {item.link}\n"
        for item in skipped
    )


def build_shopify_cart_url(site: str, items: list[BatchItem]) -> CartUrlResult:
    if not items:
        raise CartUrlError("No DLUXCA reorder links were found in the Excel file.")

    variant_quantities: defaultdict[str, int] = defaultdict(int)
    skipped: list[SkippedItem] = []
    resolved_count = 0
    for index, item in enumerate(items, start=1):
        try:
            product_url = product_json_url_from_link(site, item.link)
            product = fetch_json_url(product_url)
            allocations = select_shopify_variant_allocations(product, item)
            allocation_messages = []
            for variant, quantity, reason in allocations:
                if variant.get("available") is False:
                    raise CartUrlError(
                        f"selected unavailable variant {variant.get('id')} "
                        f"({variant.get('title') or 'Default'})"
                    )
                variant_id = str(variant.get("id") or "")
                if not variant_id:
                    raise CartUrlError("selected variant without an id")
                variant_quantities[variant_id] += quantity
                allocation_messages.append(
                    f"{quantity} -> {variant_id} ({variant.get('title') or 'Default'}, {reason})"
                )
            resolved_count += 1
            print(
                f"[{index}/{len(items)}] Resolved row {item.row_number}: "
                f"{item.quantity} x {item.name or item.sku}: "
                + "; ".join(allocation_messages),
                flush=True,
            )
        except CartUrlError as exc:
            skipped.append(
                SkippedItem(
                    row_number=item.row_number,
                    name=item.name,
                    sku=item.sku,
                    link=item.link,
                    quantity=item.quantity,
                    reason=str(exc),
                )
            )
            print(f"[{index}/{len(items)}] Skipped row {item.row_number}: {exc}", flush=True)

    cart_parts = [f"{variant_id}:{quantity}" for variant_id, quantity in variant_quantities.items()]
    if not cart_parts:
        raise CartUrlError("No unambiguous DLUXCA variants were resolved for the cart URL.")
    return CartUrlResult(
        cart_url=f"{site}/cart/{','.join(cart_parts)}",
        resolved_count=resolved_count,
        resolved_quantity=sum(variant_quantities.values()),
        skipped=skipped,
    )


def build_shopify_cart_permalink(site: str, items: list[BatchItem], skipped_path: Path) -> str:
    result = build_shopify_cart_url(site, items)
    if result.skipped:
        skipped_path.parent.mkdir(parents=True, exist_ok=True)
        skipped_path.write_text(format_skipped_items(result.skipped), encoding="utf-8")
        print(f"Skipped {len(result.skipped)} row(s). Details saved to {skipped_path}", flush=True)
    elif skipped_path.exists():
        skipped_path.unlink()
    return result.cart_url


def output_paths_for_excel(input_path: Path, output_dir: Path) -> tuple[Path, Path]:
    base_name = output_base_for_excel(input_path)
    return (
        output_dir / f"{base_name}_dluxca_cart_url.txt",
        output_dir / f"{base_name}_dluxca_cart_url_skipped.txt",
    )


def process_one_file(input_path: Path, args: argparse.Namespace, output_dir: Path) -> int:
    items = load_dluxca_items_from_excel(input_path, args)
    if args.limit is not None:
        items = items[: args.limit]

    cart_url_path, skipped_path = output_paths_for_excel(input_path, output_dir)
    cart_url = build_shopify_cart_permalink(normalize_site(args.site), items, skipped_path)
    cart_url_path.parent.mkdir(parents=True, exist_ok=True)
    cart_url_path.write_text(cart_url + "\n", encoding="utf-8")
    print(f"Wrote DLUXCA cart URL for {len(items)} row(s): {cart_url_path}")
    return len(items)


def main() -> int:
    args = parse_args()
    output_dir = Path(args.output_dir).expanduser()
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.excel:
        process_one_file(Path(args.excel).expanduser(), args, output_dir)
        return 0

    input_dir = Path(args.input_dir).expanduser()
    input_files = sorted(path for path in input_dir.glob("*.xlsx") if not path.name.startswith("~$"))
    if not input_files:
        raise CartUrlError(f"No .xlsx files found in {input_dir}.")

    total_items = 0
    for input_path in input_files:
        total_items += process_one_file(input_path, args, output_dir)
    print(f"Processed {len(input_files)} file(s), {total_items} DLUXCA reorder row(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
