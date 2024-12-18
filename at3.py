import pandas as pd
import requests
from bs4 import BeautifulSoup
import urllib.parse

def search_paddy_pallin(search_string):
    """
    Search for a product on Paddy Pallin website and check if no results are found.
    
    Args:
        search_string (str): SKU or product to search
    
    Returns:
        bool: True if no products found, False otherwise
    """
    # Encode the search string for URL
    encoded_search = urllib.parse.quote(search_string)
    
    # Construct the search URL
    url = f"https://www.paddypallin.com.au/nsearch?q={encoded_search}"
    
    try:
        # Send a request with a user agent to avoid potential blocking
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        
        # Check if request was successful
        if response.status_code != 200:
            print(f"Error fetching URL for {search_string}: HTTP {response.status_code}")
            return False
        
        # Parse the HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Look for the no results div
        no_results_div = soup.find('div', class_='nxt-nrf-container')
        
        if no_results_div:
            # Check if the div contains the "did not match any products" text
            if "did not match any products" in no_results_div.get_text():
                return True
        
        return False
    
    except requests.RequestException as e:
        print(f"Request error for {search_string}: {e}")
        return False

def process_excel_file(input_file, output_file):
    """
    Process the Excel file, search for each SKU, and output results.
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file
    """
    # Read the input Excel file
    df = pd.read_excel(input_file)
    
    # List to store results
    results = []
    
    # Iterate through each row
    for index, row in df.iterrows():
        sku = str(row['Item Code'])
        
        # Check if the SKU returns no results
        if search_paddy_pallin(sku):
            results.append({
                'SKU': sku,
                'No Results Found': True
            })
        else:
            results.append({
                'SKU': sku,
                'No Results Found': False
            })
        
        # Optional: print progress
        print(f"Processed {index + 1}/{len(df)} SKUs")
    
    # Convert results to DataFrame
    results_df = pd.DataFrame(results)
    
    # Save to output Excel file
    results_df.to_excel(output_file, index=False)
    
    print(f"Results saved to {output_file}")

# Example usage
if __name__ == "__main__":
    # input_file = './products.xlsx'  # Replace with your input file path
    input_file = './test.xlsx'  # Replace with your input file path
    output_file = 'paddy_pallin_search_results.xlsx'
    
    process_excel_file(input_file, output_file)