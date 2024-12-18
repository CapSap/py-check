import pandas as pd
import requests
from bs4 import BeautifulSoup

# Path to your Excel file
EXCEL_FILE = "products.xlsx"

# Column name in the Excel file that contains SKUs
SKU_COLUMN = "SKU"

# Base search URL (replace "${sku}" with the actual SKU)
SEARCH_URL_TEMPLATE = "https://www.paddypallin.com.au/nsearch?q={sku}"

# File to save missing SKUs
MISSING_PRODUCTS_FILE = "missing_products.xlsx"

def check_product_exists(sku):
    """
    Check if a product exists on the website by searching with the SKU.
    Returns True if the product is found, False otherwise.
    """
    # Generate the search URL for this SKU
    search_url = SEARCH_URL_TEMPLATE.format(sku=sku)
    try:
        response = requests.get(search_url, timeout=10)
        response.raise_for_status()  # Raise error for bad status codes
        soup = BeautifulSoup(response.content, 'html.parser')

        # Logic to determine if the product exists
        # Example: Look for specific elements on the search result page
        if soup.find("div", class_="no-results"):  # Replace with the actual "no results" element class
            return False
        return True
    except requests.RequestException as e:
        print(f"Error fetching URL for SKU {sku}: {e}")
        return False

def main():
    # Read the Excel file
    try:
        df = pd.read_excel(EXCEL_FILE)

        # Check if required column exists
        if SKU_COLUMN not in df.columns:
            print(f"Error: Excel file must contain a '{SKU_COLUMN}' column.")
            return

        # Prepare a list to track missing SKUs
        missing_skus = []

        # Loop through each SKU
        for index, row in df.iterrows():
            sku = row[SKU_COLUMN]
            print(f"Checking SKU: {sku}")

            if not check_product_exists(sku):
                print(f"Product not found for SKU: {sku}")
                missing_skus.append({"SKU": sku})

        # Save missing SKUs to a new Excel file
        if missing_skus:
            missing_df = pd.DataFrame(missing_skus)
            missing_df.to_excel(MISSING_PRODUCTS_FILE, index=False)
            print(f"Missing products saved to {MISSING_PRODUCTS_FILE}")
        else:
            print("All products are present.")

    except FileNotFoundError:
        print(f"Error: File {EXCEL_FILE} not found.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
