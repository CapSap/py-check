import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import urllib.parse
from datetime import datetime

start_time = datetime.now()

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
    Search for a product on Paddy Pallin website and check if a product is found.
    
    Args:
        driver (webdriver.Chrome): Selenium WebDriver
        search_string (str): SKU or product to search
    
    Returns:
        bool: True if product is found, False otherwise
    """
    # Encode the search string for URL
    encoded_search = urllib.parse.quote(search_string)
    
    # Construct the search URL
    url = f"https://www.paddypallin.com.au/nsearch?q={encoded_search}"
    
    try:
        # Navigate to the URL
        driver.get(url)
        
        # Wait for the search results container to load
        # Adjust the timeout as needed (currently set to 10 seconds)
        wait = WebDriverWait(driver, 10)
        
        try:
            # Try to find product list or page title wrapper
            # This will look for either the product list or page title to determine if a product exists
            product_locators = [
                (By.ID, "amasty-shopby-product-list"),
                (By.CLASS_NAME, "page-title-wrapper")
            ]
            
            # Check for the existence of either locator
            for locator_type, locator_value in product_locators:
                try:
                    product_element = wait.until(
                        EC.presence_of_element_located((locator_type, locator_value))
                    )
                    # If element is found, return True (product exists)
                    return True
                except:
                    # If this locator isn't found, continue to next
                    continue
            
            # If no locators are found, return False
            return False
        
        except Exception as e:
            # If any unexpected error occurs, print and return False
            print(f"Error searching for {search_string}: {e}")
            return False
        
    except Exception as e:
        print(f"Error searching for {search_string}: {e}")
        return False

def process_excel_file(input_file, output_file):
    """
    Process the Excel file, search for each SKU, and output results.
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file
    """
    # Setup WebDriver
    driver = setup_webdriver()
    
    try:
        # Read the input Excel file
        df = pd.read_excel(input_file)
        
        # List to store results
        results = []
        
        # Iterate through each row
        for index, row in df.iterrows():
            sku = str(row['Item Code'])
            
            # Check if the SKU returns product exists
            try:
                product_exists = search_paddy_pallin(driver, sku)
                row_data = row.to_dict()  # Convert the row to a dictionary
                row_data['Product Exists'] = product_exists
                results.append(row_data)                
                
                # Add a small delay between searches to reduce load on the server
                time.sleep(1)
            except Exception as e:
                print(f"Error processing SKU {sku}: {e}")
                
                # Add the error information to the row and append
                row_data = row.to_dict()
                row_data['Product Exists'] = None
                row_data['Error'] = str(e)
                
                results.append(row_data) 

            # Optional: print progress
            print(f"Processed {index + 1}/{len(df)} SKUs")
        
        # Convert results to DataFrame
        results_df = pd.DataFrame(results)
        
        # Save to output Excel file
        results_df.to_excel(output_file, index=False)
        
        print(f"Results saved to {output_file}")
    
    finally:
        # Always close the browser
        driver.quit()

# Example usage
if __name__ == "__main__":
    input_file = 'test.xlsx'  # Replace with your input file path
    output_file = 'paddy_pallin_search_results_exist_check.xlsx'
    
    process_excel_file(input_file, output_file)


    end_time = datetime.now()
    print('Duration: {}'.format(end_time - start_time))