# Jewelry Ordering App

A small Streamlit app for turning a monday.com inventory export into a DLUXCA cart
URL.

The app lets you upload one `.xlsx` file, finds rows that need to be reordered,
calculates the quantity needed, and generates a DLUXCA cart URL that can be
opened in the browser. It also groups manual review orders by store.

## Run The App

From this project folder:

```bash
source .venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Then open the local Streamlit URL, usually:

```text
http://localhost:8501
```

## How To Use It

1. Export the inventory board from monday.com as an `.xlsx` file.
2. Upload that file in the app.
3. Click `Generate cart URL`.
4. Open the generated DLUXCA cart URL.
5. Review the manual review sections and add those items manually if needed.

## Quantity Logic

The app only includes rows where `Reorder Needed` is set to `yes`.

Quantity to order is calculated as:

```text
Quantity Desired - On Floor (#) - Back Stock (#)
```

Rows with a quantity of `0` or less are ignored.

## Supported Store

The app currently generates cart URLs for DLUXCA only.

Rows for other supplier websites are shown in the manual review area, grouped by
store. The app does not log into DLUXCA or bypass Cloudflare; it uses Shopify
product data and a cart URL instead.

## Manual Review Orders

Manual review includes DLUXCA rows that could not be safely added to the cart URL
and reorder rows from other supplier stores.

Manual review rows show:

- SKU
- Name
- Quantity needed
- Product URL
- Reason skipped, for DLUXCA skipped rows

Those should be checked manually before placing the final order.
