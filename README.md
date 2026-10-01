# Bagira → Temu (GitHub Actions)

This repository processes **only the product codes in `input/export.xlsx`**. The original Temu workbook is kept intact; the script writes product rows from row 5 of its `Template` sheet and preserves the dropdown lists and other sheets.

## Run

1. Upload this package's contents to a private GitHub repository, including the two files under `input/`.
2. Open **Actions → Bagira to Temu → Run workflow**.
3. Download the `bagira-temu-results` artifact. It contains `bagira_temu.xlsx` and `review.csv`.

The script checks each linked page's article number against the export code. For search links, missing links and mismatches, it searches Bagira by the code and accepts only an exact code match. It takes **the Temu price from `export.xlsx`** where present. For the 11 rows without a price, it searches by code and converts the site's EUR price to a price excluding 20% VAT (site price / 1.20). `review.csv` records each price source. The stock is fixed at 20 per SKU.

The export column labelled `Доставчик` is mapped to the exact manufacturer spelling in the Temu dropdown for that category. No manufacturer is guessed when the dropdown lacks that supplier. Brand, legal manufacturer details and EU responsible person are separate fields; the latter stays blank.

Product titles, descriptions, photos and any published product weight come from Bagira. No default weight or package dimensions are invented. The template contains a selected set of Temu categories; 43 of the 67 export codes have a suitable category in this file. The other 24 have no sound match. The script still collects their product information in `review.csv`, but does **not** add them to the upload workbook. Other incomplete or ambiguous products are similarly excluded. Check `review.csv` before uploading, especially site price conversions, product photos, regulatory details and missing weights or shipping settings. Products with `OK` are the rows generated in the workbook.

To test a few products locally: `python scraper.py --export input/export.xlsx --template input/temu_template.xlsx --out results --limit 3`. To run all: omit `--limit`.

The explicit `CATEGORY_BY_CODE` mapping in `scraper.py` is reviewable and can be expanded after a suitable category is added to the Temu template. An absent category is not replaced with a misleading one.
