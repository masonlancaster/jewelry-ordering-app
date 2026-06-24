# Ciara Ordering App

A small Streamlit app for turning a monday.com inventory export into a DLUXCA cart
URL.

The app lets you upload one `.xlsx` file, finds the DLUXCA rows that need to be
reordered, calculates the quantity needed, and generates a cart URL that can be
opened in the browser. It also shows any skipped items so they can be reviewed
manually.

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
5. Review the skipped orders table and add those items manually if needed.

## Quantity Logic

The app only includes rows where `Reorder Needed` is set to `yes`.

Quantity to order is calculated as:

```text
Quantity Desired - On Floor (#) - Back Stock (#)
```

Rows with a quantity of `0` or less are ignored.

## Supported Store

The app currently generates cart URLs for DLUXCA only.

Rows for other supplier websites are ignored by the cart URL generator. The app
does not log into DLUXCA or bypass Cloudflare; it uses Shopify product data and a
cart URL instead.

## Skipped Orders

An item can be skipped when the app cannot safely determine the exact DLUXCA
variant to order, or when the selected variant appears unavailable.

Skipped rows show:

- SKU
- Name
- Quantity needed
- Product URL
- Reason skipped

Those should be checked manually before placing the final order.
