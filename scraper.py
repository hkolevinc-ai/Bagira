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

# Only categories actually present in the supplied Temu template. Product codes
# with no sound match are intentionally omitted and appear in review.csv.
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
}

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


def product_data(tree, url):
    for raw in tree.xpath('//script[@type="application/ld+json"]/text()'):
        try:
            obj = json.loads(raw)
        except (ValueError, TypeError):
            continue
        items = obj if isinstance(obj, list) else obj.get("@graph", [obj]) if isinstance(obj, dict) else []
        for p in items:
            if not isinstance(p, dict) or p.get("@type") != "Product":
                continue
            code = clean(p.get("sku"))
            name = clean(p.get("name"))
            description = clean(" ".join(tree.xpath('//div[contains(@class,"js-product-info")]//text()')))
            props = {clean(a.get("name")): clean(a.get("value")) for a in p.get("additionalProperty", []) if isinstance(a, dict)}
            pictures = []
            for picture in tree.xpath('//div[contains(@class,"js-prodGallery")]//img[@data-zoom or contains(@class,"gallery__big")]'):
                link = picture.get("data-zoom") or picture.get("src")
                if link:
                    pictures.append(urljoin(BASE, link))
            if not pictures:
                img = p.get("image")
                pictures = [img] if isinstance(img, str) else img if isinstance(img, list) else []
            pictures = list(dict.fromkeys(u for u in pictures if u and u.startswith("https://")))[:10]
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
            item = product_data(tree, final)
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
            p = product_data(t, final)
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", type=Path, required=True)
    ap.add_argument("--template", type=Path, required=True)
    ap.add_argument("--out", type=Path, default=Path("results"))
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    args.out.mkdir(parents=True, exist_ok=True)
    original = openpyxl.load_workbook(args.export, data_only=True, read_only=True).active
    sources = [tuple(c.value for c in r[:6]) for r in original.iter_rows(min_row=3) if r[0].value]
    if args.limit:
        sources = sources[:args.limit]
    book = openpyxl.load_workbook(args.template)
    sheet = book["Template"]
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
                 "supplier": clean(supplier), "manufacturer": "", "price_eur_ex_vat": "",
                 "price_source": "", "weight_g": "", "status": "", "detail": ""}
        try:
            item = find_product(s, code, clean(link))
            if item is None:
                raise ValueError("NO_PRODUCT: no Bagira page with exact matching article number")
            entry["product_url"] = item["url"]
            entry["bagira_name"] = item["name"]
            entry["description"] = item["description"]
            entry["image_urls"] = " | ".join(item["images"])
            if not item["images"]:
                raise ValueError("NO_IMAGE: no product image URL")
            price, source_label = price_for(source, item)
            if price <= 0:
                raise ValueError("NO_PRICE: nonpositive price")
            entry["price_eur_ex_vat"] = str(price)
            entry["price_source"] = source_label
            weight = grams(item)
            entry["weight_g"] = weight or ""
            if not category or category not in categories:
                raise ValueError("NO_CATEGORY: no suitable category in the supplied Temu template")
            if category not in manufacturers:
                manufacturers[category] = get_manufacturer_values(book, category)
            maker = SUPPLIER_TO_DROPDOWN.get(clean(supplier).upper())
            if not maker or maker not in manufacturers[category]:
                raise ValueError("NO_MANUFACTURER: supplier absent from category dropdown")
            entry["manufacturer"] = maker
            if not weight:
                raise ValueError("NO_WEIGHT: package weight needs verification")
            # Preserve all Temu-provided validations, metadata, lookup sheets and category options.
            cells = {5: category, 12: item["name"][:500], 13: code, 14: code,
                     20: item["description"][:5000], 27: item["images"][0],
                     708: item["images"][0], 719: 20, 720: float(price),
                     721: item["url"], 724: weight, 796: maker}
            for i, picture in enumerate(item["images"][1:10], start=28):
                cells[i] = picture
            for i, picture in enumerate(item["images"][1:10], start=709):
                cells[i] = picture
            for col, value in cells.items():
                sheet.cell(output_row, col, value)
            entry["status"] = "OK"
            output_row += 1
        except (requests.RequestException, ValueError) as exc:
            reason = str(exc)
            entry["status"] = reason.split(":", 1)[0] if ":" in reason else "FETCH_ERROR"
            entry["detail"] = reason
        report.append(entry)
        LOG.info("%s/%s %s %s", n, len(sources), code, entry["status"])
        time.sleep(0.6)
    with (args.out / "review.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(report[0]) if report else ["code", "status"])
        writer.writeheader()
        writer.writerows(report)
    book.save(args.out / "bagira_temu.xlsx")
    LOG.info("Saved %s rows; %s need review", output_row - 5, sum(x["status"] != "OK" for x in report))


if __name__ == "__main__":
    main()
