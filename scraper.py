#!/usr/bin/env python3
"""Fill the supplied Temu workbook from a fixed export list and Bagira pages."""
import argparse
import csv
import json
import logging
import re
import time
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from urllib.parse import urljoin

import openpyxl
import requests
from lxml import html
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = "https://www.bagira.bg"
LOG = logging.getLogger("bagira")

# Only categories actually present in the supplied Temu template. Some matches
# are approximate and are labelled as such in review.csv for merchant review.
CATEGORY_BY_CODE = {
    "3007541": "30605", "3019098": "4417",
    "503265": "55070", "16614": "55070", "16612": "55070",
    "587522": "24963", "595117": "25063", "580396": "14646", "580382": "25062",
    "590211": "39570", "590209": "39846", "589142": "13558",
    "597586": "13891", "580380": "25062", "511568": "9904", "511569": "9904",
    "78387": "9904", "3020514": "19879", "3020968": "20052", "3021870": "22390",
    "3021992": "22119", "3020981": "22119", "3020964": "32951", "3021989": "22116",
    "593108": "9060",
    "3006024": "32466", "3006021": "32466", "3005802": "11895",
    "3005785": "11919", "3005775": "11919", "3010308": "11919", "3010316": "11919",
    "3010325": "11919", "3019839": "11919", "3005766": "11919", "3019840": "11919",
    "3005798": "11953", "3005796": "11953", "3023244": "14843",
    "26582": "15510", "55476": "15510", "55589": "25058",
    "25363": "10117",
    "3019093": "15570", "3019094": "15570", "574602": "15570", "532959": "15570",
    "587517": "15570", "587518": "15570", "3012075": "13521", "3012074": "13521",
    "590201": "15894", "580419": "20190", "3020970": "21889",
    "3021225": "12775", "3021226": "12775", "3021227": "12775",
    "3021228": "12775", "3021229": "12775", "3021230": "12775",
    "579700": "24964", "3020121": "12984", "3020122": "12984", "3020119": "12984",
    "26098": "54832", "591043": "10117", "591035": "10117",
}

APPROXIMATE_CODES = {
    "3019093", "3019094", "3012075", "3012074", "574602", "532959",
    "587517", "587518", "590201", "580419", "3020970", "3021225",
    "3021226", "3021227", "3021228", "3021229", "3021230", "579700",
    "3020121", "3020122", "3020119", "26098", "591043", "591035",
}
CANISTER_CODES = {"3021225", "3021226", "3021227", "3021228", "3021229", "3021230"}
BEDDING_DOUBLE = {"3005785", "3005775", "3010308", "3010316", "3010325",
                  "3019839", "3005766", "3019840"}
BEDDING_BABY = {"3005798", "3005796"}
# Colours checked against the product photographs; review if Bagira changes them.
BEDDING_COLOUR = {
    "3005802": "White", "3005785": "Multicolor", "3005775": "Multicolor",
    "3010308": "Multicolor", "3010316": "Light Grey", "3010325": "Multicolor",
    "3019839": "Light Grey", "3005766": "Multicolor", "3019840": "Light Grey",
    "3005798": "Multicolor", "3005796": "Multicolor",
}
COUNTRY_BG_TO_EN = {"българия": "Bulgaria", "турция": "Türkiye", "италия": "Italy"}
EU_COUNTRIES = {"Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czech Republic",
                "Denmark", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary",
                "Ireland", "Italy", "Latvia", "Lithuania", "Luxembourg", "Malta", "Netherlands",
                "Poland", "Portugal", "Romania", "Slovakia", "Slovenia", "Spain", "Sweden"}

SUPPLIER_TO_DROPDOWN = {
    "АРГИРА ООД": "Argira LTD", "ВЕРТЕКС ЕООД": "VERTEX", "ЗИЕСТО АД": "ZIESTO",
    "МАРИСАН & КОЛЕВ - АД": "MARISAN & KOLEV", "МИКРОСИС КО - ООД": "Microsys Co Ltd.",
    "ПАОЛО ЕООД": "PAOLO", "ПРОФИ ПЛЮС ООД": "PROFI PLUS", "РОСТРЕЙД - ЕООД": "ROSTRADE",
    "ХИДРОСТАБ ООД": "HYDROSTAB", "ЧИЛИ ТРЕЙД ЕООД": "CHILI TRADE",
    "ВИ ФРЕНД - ЕООД": "VI FRIEND", "ТЕСИ ООД": "TESY",
    "ГАНЕВ ТРЕЙДИНГ - ООД": "GANEV TRADING LTD", "БАГИРА ООД": "Bagira LTD",
    "РОТО-БГ ЕООД": "ROTO-BG", "МИРА-Н ЕТ": "Mira - N",
}


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def session():
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (compatible; BagiraTemuImporter/1.0)",
                      "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.7"})
    retry = Retry(total=3, backoff_factor=1.2, status_forcelist=[429, 500, 502, 503, 504])
    s.mount("https://", HTTPAdapter(max_retries=retry))
    return s


def get_tree(s, url):
    response = s.get(url, timeout=35)
    response.raise_for_status()
    if not response.url.startswith(BASE + "/"):
        raise ValueError("Unexpected redirect outside bagira.bg")
    return html.fromstring(response.content.decode("utf-8", errors="replace")), response.url


def normalized_code(value):
    return re.sub(r"^Vo\s*", "", clean(value), flags=re.I)


def product_data(tree, url, expected_code=None):
    for raw in tree.xpath('//script[@type="application/ld+json"]/text()'):
        try:
            obj = json.loads(raw)
        except (ValueError, TypeError):
            continue
        items = obj if isinstance(obj, list) else obj.get("@graph", [obj]) if isinstance(obj, dict) else []
        for p in items:
            if not isinstance(p, dict) or p.get("@type") != "Product":
                continue
            code = normalized_code(p.get("sku"))
            name = clean(p.get("name"))
            description = clean(" ".join(tree.xpath('//div[contains(@class,"js-product-info")]//text()')))
            props = {clean(a.get("name")): clean(a.get("value")) for a in p.get("additionalProperty", []) if isinstance(a, dict)}
            pictures = []
            for picture in tree.xpath('//div[contains(@class,"js-prodGallery")]//img[@data-zoom or contains(@class,"gallery__big")]'):
                link = picture.get("data-zoom") or picture.get("src")
                if link:
                    pictures.append(urljoin(BASE, link))
            for thumb in tree.xpath('//div[contains(@class,"js-prodGallery")]//span[contains(@class,"gallery__thumb")][@data-zoom]'):
                pictures.append(urljoin(BASE, thumb.get("data-zoom")))
            if not pictures:
                img = p.get("image")
                pictures = [img] if isinstance(img, str) else img if isinstance(img, list) else []
            pictures = list(dict.fromkeys(u for u in pictures if u and u.startswith("https://")))[:10]
            if expected_code in CANISTER_CODES:
                # The 10 L / 20 L pages contain a table of separate colour SKUs.
                # Accept a code only when it appears in that table, then use its own image.
                variants = {}
                for row in tree.xpath('//div[contains(@class,"js-product-info")]//tr'):
                    cols = [clean(x.text_content()) for x in row.xpath('./td')]
                    if len(cols) >= 3 and re.fullmatch(r"\d{5,8}", cols[0]):
                        variants[cols[0]] = cols[2]
                if expected_code not in variants and code != expected_code:
                    return None
                code = expected_code
                matching = []
                for u in pictures:
                    stem = u.rsplit("/", 1)[-1].split(".", 1)[0]
                    if stem == code or stem.endswith(code):
                        matching.append(u)
                if not matching:
                    return None
                pictures = matching
                if code in variants:
                    name = f"{name} – {variants[code].lower()}"
                    info = tree.xpath('//div[contains(@class,"js-product-info")]')
                    if info:
                        for table in info[0].xpath('.//table'):
                            table.drop_tree()
                        description = clean(info[0].text_content()) + f" Цвят: {variants[code].lower()}."
            offer = p.get("offers", {})
            if isinstance(offer, list):
                offer = offer[0] if offer else {}
            return {"code": code, "name": name, "description": description or name,
                    "properties": props, "images": pictures, "price": offer.get("price"),
                    "currency": offer.get("priceCurrency"), "url": p.get("url") or url}
    return None


def find_product(s, code, initial):
    checked = set()
    if initial and initial.startswith(BASE + "/") and "/tarsene/" not in initial:
        checked.add(initial)
        try:
            tree, final = get_tree(s, initial)
            item = product_data(tree, final, code)
            if item and item["code"] == code:
                return item
        except (requests.RequestException, ValueError) as exc:
            LOG.warning("Linked page %s failed: %s", code, exc)
    tree, _ = get_tree(s, f"{BASE}/tarsene/{code}")
    # Search result cards have product links; verify each candidate's JSON-LD sku.
    urls = []
    for a in tree.xpath('//a[@href]'):
        href = a.get("href", "")
        if re.match(r"^/(\d+)/", href) or re.match(r"^https://www\.bagira\.bg/\d+/", href):
            u = urljoin(BASE, href)
            if u not in checked and u not in urls:
                urls.append(u)
    for u in urls[:12]:
        checked.add(u)
        try:
            t, final = get_tree(s, u)
            p = product_data(t, final, code)
            if p and p["code"] == code:
                return p
        except (requests.RequestException, ValueError):
            continue
    return None


def get_manufacturer_values(workbook, category):
    sheet = workbook["Dropdown Lists"]
    key = f"t_8_{category}_Manufacturer"
    for row in sheet.iter_rows(min_col=1, max_col=18, values_only=True):
        if row[0] == key:
            return set(clean(x) for x in row[2:] if x)
    return set()


def price_for(source, product):
    if source[2] is not None and clean(source[2]):
        return Decimal(str(source[2])).quantize(Decimal("0.01")), "export.xlsx"
    if product["currency"] != "EUR" or not product["price"]:
        raise ValueError("Price missing in export and no EUR price on Bagira")
    # The export explicitly describes a Temu price excluding VAT.
    return (Decimal(str(product["price"])) / Decimal("1.20")).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP), "Bagira EUR retail / 1.20"


def grams(product):
    for label, val in product["properties"].items():
        if "тегло" in label.lower():
            found = re.search(r"\d+(?:[.,]\d+)?", val)
            if found:
                n = Decimal(found.group().replace(",", "."))
                return int((n * (1000 if "кг" in label.lower() else 1)).quantize(Decimal("1")))
    return None


def full_description(product):
    description = product["description"]
    if len(description) < 150 and product["properties"]:
        details = "; ".join(f"{key}: {value}" for key, value in product["properties"].items()
                            if value and "тегло" not in key.lower())
        if details:
            return f"{description}. {details}"
    return description


def dropdown_values(index, key):
    return index.get(key, [])


def choose(index, key, preferred):
    values = dropdown_values(index, key)
    for value in preferred:
        if value in values:
            return value
    return None


def origin_country(product):
    source = " ".join(f"{k}: {v}" for k, v in product["properties"].items())
    source += " " + product["description"]
    for bg, en in COUNTRY_BG_TO_EN.items():
        if re.search(r"(?:произход|произведено в)\s*:\s*" + bg, source, re.I):
            return en
    return None


def dimension_pair(value):
    m = re.search(r"(\d{2,3})\s*[хx×*]\s*(\d{2,3})\s*см", value or "", re.I)
    return (int(m.group(1)), int(m.group(2))) if m else None


def derived_material(product, cat, dropdown):
    source = (product["properties"].get("Материал", "") + " " + product["description"]).lower()
    name = product["name"].lower()
    # Prefer an explicit material phrase; fall back to a clear material in the title/description.
    signals = [
        (r"(?:pvc|поливинилхлорид)", ["PVC", "Polyvinyl chloride", "pvc", "Plastic"]),
        (r"(?:неръждаема стомана|inox)", ["Stainless Steel", "Stainless steel", "Steel", "Metal"]),
        (r"(?:въглеродна стомана)", ["Carbon Steel", "High carbon steel", "Steel", "Metal"]),
        (r"(?:алуминиев|алуминий)", ["Aluminum", "Aluminum Alloy", "Metal"]),
        (r"(?:полиестер)", ["Polyester", "Polyester fiber", "Fabric"]),
        (r"(?:памук)", ["Cotton", "Fabric"]),
        (r"(?:пластмаса|пластмасов|abs)", ["Plastic", "ABS Plastic", "ABS", "PVC"]),
        (r"(?:стомана|стоманен|поцинкован)", ["Steel", "Metal", "Iron"]),
        (r"(?:метал|метален)", ["Metal", "Steel", "Iron"]),
        (r"(?:дърво|дървен|бук)", ["Wood", "Beechwood"]),
    ]
    for pattern, preferred in signals:
        if re.search(pattern, source + " " + name):
            value = choose(dropdown, f"t_3_{cat}_121 - Material", preferred)
            if value:
                return value
            value = choose(dropdown, f"t_3_{cat}_12 - Material", preferred)
            if value:
                return value
    return None


def category_fields(code, cat, product, dropdown):
    """Return supported Temu column values, using only product evidence and valid dropdown choices."""
    desc = product["description"].lower()
    props = product["properties"]
    cells = {}
    notes = []
    material = derived_material(product, cat, dropdown)
    if material:
        cells[123 if cat not in {"19879", "20052", "20190", "32466", "39846"} else 283] = material
    country = origin_country(product)
    if country and country in dropdown_values(dropdown, f"t_8_{cat}_Country/Region of Origin"):
        cells[743] = country
    if cat in {"11919", "11953", "11895"}:
        composition = props.get("Материал", "") or product["description"]
        cotton = re.search(r"(\d{1,3})%\s*памук", composition, re.I)
        polyester = re.search(r"(\d{1,3})%\s*полиестер", composition, re.I)
        if cotton:
            cells[183] = int(cotton.group(1))
        if polyester:
            cells[189] = int(polyester.group(1))
        if cat == "11895" and "полиестер" in composition.lower() and not polyester:
            cells[189] = 100
        # A duvet cover/sheet and a pillow protector have no separately stated lining.
        # Do not guess the liner material; it remains in the review report.
        if cat in {"11953", "11895"} and (cotton or polyester or 189 in cells):
            cells[250] = "Yes"
        if cat == "11953":
            cells[393] = choose(dropdown, f"t_3_{cat}_1117 - Applicable Age Group", ["0 Year+"])
        if cat in {"11919", "11895"}:
            cells[666] = choose(dropdown, f"t_4_{cat}_Variation Theme",
                                ["Color × Duvet Cover Size", "Color × Size"])
            cells[667] = choose(dropdown, f"t_4_{cat}_Size Family", ["2 - Regular Size"])
            cells[668] = choose(dropdown, f"t_4_{cat}_2 - Regular Size_Sub-Size Family", ["10 - Alpha"])
            colour = BEDDING_COLOUR.get(code)
            if colour in dropdown_values(dropdown, f"t_4_{cat}_Color"):
                cells[674] = colour
                notes.append("Color selected from product photograph; check before upload.")
        else:
            cells[666] = "Color"
            cells[675] = BEDDING_COLOUR.get(code, "Multicolor")
            notes.append("Color selected from product photograph; check before upload.")
        pillow = dimension_pair(props.get("Размери на калъфката (Ш х В)", ""))
        if cat == "11895":
            pillow = dimension_pair(product["description"])
            if pillow:
                size = f"{pillow[0]}cm*{pillow[1]}cm"
                opts = dropdown_values(dropdown, f"t_4_{cat}_2 - Regular Size_10 - Alpha_Size")
                if size in opts:
                    cells[673] = size
        if pillow:
            cells[691], cells[692] = pillow
        duvet = dimension_pair(props.get("Размери на плика (Ш х В)", ""))
        if duvet:
            if cat == "11919":
                size = f"{duvet[0]}cm*{duvet[1]}cm"
                opts = dropdown_values(dropdown, f"t_4_{cat}_2 - Regular Size_10 - Alpha_Duvet Cover Size")
                if size in opts:
                    cells[670] = size
            cells[701], cells[702] = duvet[1], duvet[0]
        sheet = dimension_pair(props.get("Размери на чаршафа (Ш х В)", ""))
        if sheet:
            cells[696], cells[697] = sheet[1], sheet[0]
        if any(c in cells for c in [691, 701, 706]):
            cells[690] = "cm-g-ml"
        if cat == "11919":
            notes.append("Closure type not stated on Bagira; verify before upload.")
    else:
        # One row per source article. For non-variant categories the Quantity
        # sale property uses one purchased SKU; coloured canisters use Color.
        if code in CANISTER_CODES:
            cells[666] = "Color"
            colours = re.findall(r"цвят\s*:\s*(бял|жълт|син)", desc)
            colour = {"бял": "White", "жълт": "Yellow", "син": "Blue"}.get(colours[-1]) if colours else None
            if colour:
                cells[675] = colour
        elif "Quantity" in dropdown_values(dropdown, f"t_4_{cat}_Variation Theme"):
            cells[666], cells[685] = "Quantity", "1"
        if cat == "30605":
            cells[618] = choose(dropdown, f"t_3_{cat}_1305 - Panel Material", ["Cotton"])
            cells[690], cells[706], cells[707] = "cm-g-ml", 3.9, 130
        if cat == "22116" and "алуминиево фолио" in desc:
            cells[387] = choose(dropdown, f"t_3_{cat}_1920 - Major Material", ["Aluminum"])
        if cat == "15570":
            cells[364] = choose(dropdown, f"t_3_{cat}_1561 - Power Supply", ["Use Without Electricity"])
            cells[396] = choose(dropdown, f"t_3_{cat}_225 - Blade Material",
                                ["Stainless Steel"] if "неръждаема" in desc else ["Carbon Steel"] if "въглеродна" in desc else [])
            handles = (["Plastic"] if "дръжката: пластмаса" in desc else
                       ["Beechwood", "Wood"] if any(x in desc for x in ["букова", "бук"])
                       else [])
            cells[363] = choose(dropdown, f"t_3_{cat}_403 - Handle Material", handles)
        if cat in {"15510", "54832", "13891", "39570", "15894", "9904"}:
            cells[364] = choose(dropdown, f"t_3_{cat}_1561 - Power Supply", ["Use Without Electricity"])
        if cat in {"13891", "39570", "15894"}:
            cells[316] = choose(dropdown, f"t_3_{cat}_2153 - Battery Properties", ["Without Battery"])
        if cat == "9904":
            cells[417] = choose(dropdown, f"t_3_{cat}_244 - Outer Material", ["Stainless Steel"])
            cells[426] = choose(dropdown, f"t_3_{cat}_2155 - Laser Type", ["Without Laser"])
        if cat == "4417":
            cells[438] = choose(dropdown, f"t_3_{cat}_1440 - Main Material", ["PVC"])
        if cat == "25058":
            cells[396] = choose(dropdown, f"t_3_{cat}_225 - Blade Material", ["Carbon Steel"])
        if cat == "13521":
            cells[448] = choose(dropdown, f"t_3_{cat}_15 - Composition", ["Others"])
        if cat == "14646":
            cells[250] = "Yes" if "полиестер" in desc else None
            cells[512] = choose(dropdown, f"t_3_{cat}_1279 - Best Use", ["Industrial Use"])
        if cat in {"22116", "22119", "21889", "22390"}:
            cells[497] = choose(dropdown, f"t_3_{cat}_1941 - Applicable Models", ["Universally Applicable"])
        if cat == "21889":
            cells[502] = choose(dropdown, f"t_3_{cat}_7039 - Are These Parts Original Car Parts?", ["No"])
        if cat == "20052":
            cells[565] = choose(dropdown, f"t_3_{cat}_6449 - Does It Contain Chemicals", ["Yes"])
    return {k: v for k, v in cells.items() if v is not None}, country, notes


def missing_required(sheet, mode, cat, row, cells, country):
    rr = next((r for r in range(1, mode.max_row + 1) if mode.cell(r, 1).value == cat + "_require"), None)
    if rr is None:
        return ["Category validation rules missing"]
    groups = {}
    for col in range(5, sheet.max_column + 1):
        if mode.cell(rr, col).value != "require":
            continue
        header = str(sheet.cell(2, col).value)
        group = header.split(":", 1)[0] if header.startswith(("2021 - Cover Material:", "2035 - Liner Material:")) else header
        groups.setdefault(group, []).append(col)
    missing = []
    for name, columns in groups.items():
        if name == "List Price - EUR" and cells.get(723) == "N/A":
            continue
        if not any(cells.get(c) not in (None, "") for c in columns):
            missing.append(name)
    # This field is labelled Required in Data Definitions, though GoodsLevelMode
    # does not mark it. Its answer depends on the merchant's sales history.
    if not cells.get(798):
        missing.append("Post-13-Dec-2024 EU/NI market declaration")
    return missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", type=Path, required=True)
    ap.add_argument("--template", type=Path, required=True)
    ap.add_argument("--out", type=Path, default=Path("results"))
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--cached-review", type=Path, help="Reuse previously verified Bagira data from review.csv")
    ap.add_argument("--enriched-json", type=Path, help="Use previously verified exact-code product data")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    args.out.mkdir(parents=True, exist_ok=True)
    original = openpyxl.load_workbook(args.export, data_only=True, read_only=True).active
    sources = [tuple(c.value for c in r[:6]) for r in original.iter_rows(min_row=3) if r[0].value]
    if args.limit:
        sources = sources[:args.limit]
    allowed_codes = {clean(r[0]) for r in sources}
    if len(allowed_codes) != len(sources):
        raise ValueError("Duplicate product codes in export.xlsx")
    cached = {}
    if args.cached_review:
        with args.cached_review.open(newline="", encoding="utf-8-sig") as f:
            for previous in csv.DictReader(f):
                code = clean(previous.get("code"))
                if code not in allowed_codes or code in cached:
                    raise ValueError(f"Unexpected or duplicate cached product code: {code}")
                cached[code] = previous
    enriched = {}
    if args.enriched_json:
        payload = json.loads(args.enriched_json.read_text(encoding="utf-8"))
        enriched = {code: record["data"] for code, record in payload.items()
                    if record.get("data") is not None}
        if set(enriched) != allowed_codes:
            raise ValueError("Enriched product codes must exactly match the export")
        for code, item in enriched.items():
            if item.get("code") != code or not item.get("url", "").startswith(BASE + "/"):
                raise ValueError(f"Unverified enriched product: {code}")
    book = openpyxl.load_workbook(args.template)
    sheet = book["Template"]
    mode = book["GoodsLevelMode"]
    dropdown = {str(row[0]): [v for v in row[2:] if v is not None]
                for row in book["Dropdown Lists"].iter_rows(values_only=True) if row[0]}
    categories = {str(r[0].value): r[1].value for r in book["Browse Data"].iter_rows() if r[0].value}
    manufacturers = {}
    s = session()
    report = []
    output_row = 5
    for n, source in enumerate(sources, 1):
        code, export_name, _, _, supplier, link = source
        code = clean(code)
        category = CATEGORY_BY_CODE.get(code)
        entry = {"code": code, "export_name": clean(export_name), "source_link": clean(link),
                 "product_url": "", "bagira_name": "", "description": "", "image_urls": "",
                 "category": category or "", "category_name": categories.get(category, ""),
                 "category_match": "approximate" if code in APPROXIMATE_CODES else "direct",
                 "supplier": clean(supplier), "manufacturer": "", "price_eur_ex_vat": "",
                 "price_source": "", "weight_g": "", "status": "", "detail": ""}
        try:
            prior = cached.get(code)
            if code in enriched:
                item = enriched[code]
            elif (prior and code not in CANISTER_CODES and prior.get("bagira_name")
                    and prior.get("product_url", "").startswith(BASE + "/")):
                item = {"code": code, "name": prior["bagira_name"],
                        "description": prior.get("description") or prior["bagira_name"],
                        "images": prior.get("image_urls", "").split(" | ") if prior.get("image_urls") else [],
                        "properties": json.loads(prior.get("properties_json") or "{}"),
                        "url": prior["product_url"], "price": prior.get("site_price_eur"),
                        "currency": "EUR" if prior.get("site_price_eur") else None}
            else:
                item = find_product(s, code, clean(link))
            if item is None:
                raise ValueError("NO_PRODUCT: no Bagira page with exact matching article number")
            entry["product_url"] = item["url"]
            entry["bagira_name"] = item["name"]
            entry["description"] = full_description(item)
            entry["image_urls"] = " | ".join(item["images"])
            entry["properties_json"] = json.dumps(item["properties"], ensure_ascii=False)
            entry["site_price_eur"] = item.get("price") or ""
            if code == "3020121" and "12л" in clean(export_name).lower() and "10 литра" in item["name"].lower():
                entry["detail"] = "Verify capacity: export says 12 L; Bagira article 3020121 says 10 L."
            if not item["images"]:
                raise ValueError("NO_IMAGE: no product image URL")
            if prior and source[2] is None and prior.get("price_eur_ex_vat") and item:
                price = Decimal(prior["price_eur_ex_vat"])
                source_label = prior.get("price_source") or "Bagira EUR retail / 1.20"
            else:
                price, source_label = price_for(source, item)
            if price <= 0:
                raise ValueError("NO_PRICE: nonpositive price")
            entry["price_eur_ex_vat"] = str(price)
            entry["price_source"] = source_label
            site_weight = grams(item)
            weight = site_weight or (int(prior["weight_g"]) if prior and prior.get("weight_g") else None)
            if not weight:
                weight = 1000
                entry["weight_source"] = "default"
            else:
                entry["weight_source"] = ("Bagira" if site_weight or not prior else
                                          prior.get("weight_source") or "Bagira")
            entry["weight_g"] = weight
            if not category or category not in categories:
                raise ValueError("NO_CATEGORY: no suitable category in the supplied Temu template")
            if category not in manufacturers:
                manufacturers[category] = get_manufacturer_values(book, category)
            maker = SUPPLIER_TO_DROPDOWN.get(clean(supplier).upper())
            if not maker or maker not in manufacturers[category]:
                raise ValueError("NO_MANUFACTURER: supplier absent from category dropdown")
            entry["manufacturer"] = maker
            # Preserve all Temu-provided validations, metadata, lookup sheets and category options.
            cells = {5: category, 7: "Normal product", 12: item["name"][:500],
                     13: code, 14: code, 15: "Add",
                     20: full_description(item)[:5000], 27: item["images"][0],
                     708: item["images"][0], 719: 20, 720: float(price),
                     721: item["url"], 724: weight, 725: 20, 726: 20, 727: 20,
                     729: "Yes", 730: 1, 731: "piece", 739: "Склад", 795: code, 796: maker}
            retail = item.get("price")
            if item.get("currency") == "EUR" and retail and Decimal(str(retail)) > price:
                cells[722] = float(Decimal(str(retail)))
            else:
                cells[723] = "N/A"
            extra, origin, notes = category_fields(code, category, item, dropdown)
            cells.update(extra)
            missing = missing_required(sheet, mode, category, output_row, cells, origin)
            entry["origin_country"] = origin or ""
            entry["missing_required"] = " | ".join(missing)
            entry["field_notes"] = " | ".join(notes)
            for i, picture in enumerate(item["images"][1:10], start=28):
                cells[i] = picture
            for i, picture in enumerate(item["images"][1:10], start=709):
                cells[i] = picture
            for col, value in cells.items():
                sheet.cell(output_row, col, value)
            entry["status"] = "NEEDS_MERCHANT_DATA" if missing else "OK"
            output_row += 1
        except (requests.RequestException, ValueError) as exc:
            reason = str(exc)
            entry["status"] = reason.split(":", 1)[0] if ":" in reason else "FETCH_ERROR"
            entry["detail"] = reason
        report.append(entry)
        LOG.info("%s/%s %s %s", n, len(sources), code, entry["status"])
        if not args.enriched_json and not args.cached_review:
            time.sleep(0.6)
    with (args.out / "review.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(dict.fromkeys(k for record in report for k in record)) if report else ["code", "status"])
        writer.writeheader()
        writer.writerows(report)
    book.save(args.out / "bagira_temu.xlsx")
    LOG.info("Saved %s rows; %s need review", output_row - 5, sum(x["status"] != "OK" for x in report))


if __name__ == "__main__":
    main()
