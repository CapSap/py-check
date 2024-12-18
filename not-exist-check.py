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
    chrome_options.add_argument("--headless")
    
    # Setup the WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    return driver

def search_paddy_pallin(driver, search_string):
    """
    Search for a product on Paddy Pallin website and check if no results are found.
    
    Args:
        driver (webdriver.Chrome): Selenium WebDriver
        search_string (str): SKU or product to search
    
    Returns:
        bool: True if no products found, False otherwise
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
            # Try to find the no results container
            no_results_container = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, 'nxt-nrf-container'))
            )
            
            # Check if the text indicates no results
            if "did not match any products" in no_results_container.text:
                return True
            
        except:
            # If the no results container is not found, assume products exist
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
            
            # Check if the SKU returns no results
            try:
                no_results = search_paddy_pallin(driver, sku)
                row_data = row.to_dict()  # Convert the row to a dictionary
                row_data['No Results Found'] = no_results
                results.append(row_data)                
                
                # Add a small delay between searches to reduce load on the server
                time.sleep(1)
            except Exception as e:
                print(f"Error processing SKU {sku}: {e}")
                
                # Add the error information to the row and append
                row_data = row.to_dict()
                row_data['No Results Found'] = None
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
    # input_file = 'test.xlsx'  
    input_file = 'products.xlsx' 
    output_file = 'paddy_pallin_search_results.xlsx'
    
    process_excel_file(input_file, output_file)
    end_time = datetime.now()
    print('Duration: {}'.format(end_time - start_time))