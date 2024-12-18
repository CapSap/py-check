import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import concurrent.futures
import urllib.parse

def setup_webdriver():
    """
    Set up Chrome WebDriver with options for more reliable scraping.
    
    Returns:
        webdriver.Chrome: Configured Chrome WebDriver
    """
    # Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # Uncomment the next line to run in headless mode
    # chrome_options.add_argument("--headless")
    
    # Setup the WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    return driver

def search_paddy_pallin(driver, search_string):
    """
    Search for a product on Paddy Pallin website and check for results.
    
    Args:
        driver (webdriver.Chrome): Selenium WebDriver
        search_string (str): SKU or product to search
    
    Returns:
        dict: Search result information
    """
    # Encode the search string for URL
    encoded_search = urllib.parse.quote(search_string)
    
    # Construct the search URL
    url = f"https://www.paddypallin.com.au/nsearch?q={encoded_search}"
    
    try:
        # Navigate to the URL
        driver.get(url)
        
        # Wait for the search results container to load
        wait = WebDriverWait(driver, 10)
        
        try:
            # Look for the product list container
            product_container = wait.until(
                EC.presence_of_element_located((By.ID, 'amasty-shopby-product-list'))
            )
            
            # Check for product items
            product_items = product_container.find_elements(By.CLASS_NAME, 'product-item')
            
            # If product items exist, it's a valid result
            if product_items:
                return {
                    'Item Code': search_string,
                    'No Results Found': False,
                    'Product Count': len(product_items)
                }
            
            # If no product items, check for no results message
            no_results_container = driver.find_elements(By.CLASS_NAME, 'nxt-nrf-container')
            
            if no_results_container and "did not match any products" in no_results_container[0].text:
                return {
                    'Item Code': search_string,
                    'No Results Found': True,
                    'Product Count': 0
                }
            
            # If neither condition is met, return uncertain result
            return {
                'Item Code': search_string,
                'No Results Found': None,
                'Product Count': 0
            }
        
        except Exception as e:
            print(f"Error processing {search_string}: {e}")
            return {
                'Item Code': search_string,
                'No Results Found': None,
                'Error': str(e)
            }
        
    except Exception as e:
        print(f"Error searching for {search_string}: {e}")
        return {
            'Item Code': search_string,
            'No Results Found': None,
            'Error': str(e)
        }

def process_excel_file(input_file, output_file, max_workers=5):
    """
    Process the Excel file using concurrent searches.
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file
        max_workers (int): Number of concurrent searches
    """
    # Read the input Excel file
    df = pd.read_excel(input_file)
    
    # List to store results
    results = []
    
    # Setup WebDrivers
    drivers = [setup_webdriver() for _ in range(max_workers)]
    
    try:
        # Use ThreadPoolExecutor for concurrent searches
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Create a list to track futures
            future_to_sku = {
                executor.submit(search_paddy_pallin, drivers[i % max_workers], str(row['Item Code'])): 
                str(row['Item Code']) 
                for i, row in df.iterrows()
            }
            
            # Collect results as they complete
            for future in concurrent.futures.as_completed(future_to_sku):
                try:
                    result = future.result()
                    results.append(result)
                    
                    # Print progress
                    print(f"Processed SKU: {result['Item Code']}")
                
                except Exception as e:
                    print(f"Unexpected error: {e}")
        
        # Convert results to DataFrame
        results_df = pd.DataFrame(results)
        
        # Save to output Excel file
        results_df.to_excel(output_file, index=False)
        
        print(f"Results saved to {output_file}")
    
    finally:
        # Always close all browsers
        for driver in drivers:
            driver.quit()

# Example usage
if __name__ == "__main__":
    # input_file = './products.xlsx'  # Replace with your input file path
    input_file = './test.xlsx'  # Replace with your input file path
    output_file = 'paddy_pallin_search_results.xlsx'
    
    # Process with 5 concurrent searches
    process_excel_file(input_file, output_file, max_workers=5)