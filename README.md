# Bagira → Temu (GitHub Actions)

This repository processes **only the product codes in `input/export.xlsx`**. The original Temu workbook is kept intact; the script writes product rows from row 5 of its `Template` sheet and preserves the dropdown lists and other sheets.

## Run

1. Upload this package's contents to a private GitHub repository, including the two files under `input/`.
2. Open **Actions → Bagira to Temu → Run workflow**.
3. Download the `bagira-temu-results` artifact. It contains `bagira_temu.xlsx` and `review.csv`.

The script checks each linked page's article number against the export code. For search links, missing links and mismatches, it searches Bagira by the code and accepts only an exact code match. It takes **the Temu price from `export.xlsx`** where present. For the 11 rows without a price, it searches by code and converts the site's EUR price to a price excluding 20% VAT (site price / 1.20). `review.csv` records each price source. The stock is fixed at 20 per SKU.

The export column labelled `Доставчик` is mapped to the exact manufacturer spelling in the Temu dropdown for that category. No manufacturer is guessed when the dropdown lacks that supplier. Brand, legal manufacturer details and EU responsible person are separate fields.

Product titles, descriptions, photos and any published product weight come from Bagira. If weight is unavailable, the script enters **1 kg**. It enters **20 × 20 × 20 cm** for length, width and height because package dimensions are unavailable from this source. These values should be checked against the actual packed items before upload. The template contains a selected set of Temu categories: 43 codes have closer matches, while 24 use the nearest available category at the merchant's request. These are marked `approximate` in `review.csv` and should be checked before upload. All 67 export codes are processed; rows with unresolved product identity or missing images are excluded. `Vo` prefixes and the six separate colour SKUs on the Bagira canister pages are verified against Bagira article codes.

The scraper fills product type, update action, the appropriate variation theme, available category attributes, published textile composition and product sizes, product identification (source article code), packaging fields, and the template's `Склад` shipping option. It uses the site's EUR retail price as List Price only when it exceeds the base Temu price; otherwise it enters `N/A` for List Price. Photo-based bedding colours are marked for review. The `missing_required` column in `review.csv` checks category rules in `GoodsLevelMode` and adds the EU/NI post-2024 market declaration marked Required in `Data Definitions`. `NEEDS_MERCHANT_DATA` rows are still included in the workbook; their missing fields need merchant confirmation before upload. In particular, a supplier name or country of origin does not establish the EU responsible person's details, and the product page does not establish whether it has been placed on the EU/NI market since December 13, 2024. Do not invent these values. Check `review.csv` before uploading, especially site price conversions, category choices, product photos, regulatory details and shipping settings.

To test a few products locally: `python scraper.py --export input/export.xlsx --template input/temu_template.xlsx --out results --limit 3`. To run all: omit `--limit`.

For a quick repeat run using data from an earlier artifact: add `--cached-review path/to/review.csv`. The script still validates every cached code against the current export list. Omit this option to scrape Bagira again. `--enriched-json` is an optional internal cache of verified exact-code Bagira product records for reproducible local checks; it is not required for GitHub Actions.

**Capacity discrepancy:** export code `3020121` says 12 L, while Bagira's page for the same article code says 10 L. Its row follows Bagira's title and description; verify this item before upload. The discrepancy is also noted in `review.csv`.

The explicit `CATEGORY_BY_CODE` mapping in `scraper.py` is reviewable and can be expanded after a suitable category is added to the Temu template. An absent category is not replaced with a misleading one.

The merchant requests blank Country/Region of Origin, EU Responsible person, and post-2024 EU/NI declaration columns. Published origin is retained in `review.csv` as `source_origin_country`. Textile sets use the N/A liner composition field where no separate liner is specified; the polyester pillow protector uses its stated material. The eight double sets have flat sheets, so their fitted-sheet fields remain empty and are marked inapplicable in `review.csv`. Closure type remains unfilled where Bagira does not state it.
