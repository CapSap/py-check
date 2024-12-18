import pandas as pd
from bs4 import BeautifulSoup

def check_if_product_not_found(html_content, sku):
    """
    Check if a product is not found based on HTML content and SKU.
    
    Args:
        html_content (str): The HTML content as a string.
        sku (str): The SKU to check.
    
    Returns:
        bool: True if the product is not found, False otherwise.
    """
    # Parse the HTML content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Look for the "not found" message
    not_found_element = soup.find(id="nxt-nrf")
    if not_found_element and f"Your search - {sku} - did not match any products" in not_found_element.text:
        return True
    return False

def process_skus_from_excel(excel_file, html_contents, output_file):
    """
    Process SKUs from an Excel file and check if products are not found.
    
    Args:
        excel_file (str): Path to the Excel file with SKUs.
        html_contents (dict): A dictionary with SKUs as keys and HTML content as values.
        output_file (str): Path to save the Excel file with "not found" SKUs.
    """
    # Read SKUs from the Excel file
    df = pd.read_excel(excel_file)
    
    # Ensure there's a column named "Item Code"
    if 'Item Code' not in df.columns:
        raise ValueError("Excel file must have a column named 'Item Code'.")
    
    # List to store SKUs not found
    not_found_skus = []
    
    # Check each SKU
    for sku in df['Item Code']:
        html_content = html_contents.get(sku, "")  # Get HTML for the SKU
        if check_if_product_not_found(html_content, sku):
            not_found_skus.append(sku)
    
    # Save the not-found SKUs to an Excel file
    not_found_df = pd.DataFrame({'Not Found SKUs': not_found_skus})
    not_found_df.to_excel(output_file, index=False)
    print(f"Saved not-found SKUs to {output_file}")

# Example usage
if __name__ == "__main__":
    # Example HTML contents for testing (replace with actual HTML fetched for SKUs)
    html_contents = {
        "50931406740S": '<div id="nxt-nrf"> Your search - <strong>50931406740S</strong> - did not match any products<br><br></div>',
        "51213403473OS": '<div id="amasty-shopby-product-list">...</div>',
    }
    
    # Path to the input Excel file
    input_excel = "products.xlsx"  # Replace with the actual file path
    
    # Path to the output Excel file
    output_excel = "not_found_skus.xlsx"
    
    # Process the SKUs
    process_skus_from_excel(input_excel, html_contents, output_excel)
