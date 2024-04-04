import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

# Function to initialize the WebDriver
def initialize_driver():
    print("Initializing WebDriver...")
    try:
        # Set Chrome options
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--headless")  # Run Chrome in headless mode
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Create a new Chrome window
        driver = webdriver.Chrome(options=chrome_options)
        print("WebDriver initialized successfully.")
        return driver
    except Exception as e:
        print(f"Error initializing WebDriver: {e}")
        return None

# Function to scrape market cap, share price, trailing P/E, Price/Book (mrq), beta, 52 Week High, 52 Week Low, 50-Day Moving Average, Enterprise Value, sector, and full-time employees from the statistics and profile tabs for a given company symbol
def scrape_company_data(driver, symbol):
    try:
        symbol_with_extension = symbol + ".NS"
        
        # Scraping market cap, share price, trailing P/E, Price/Book (mrq), beta, 52 Week High, 52 Week Low, 50-Day Moving Average, and Enterprise Value from statistics tab
        print(f"Scraping data for symbol {symbol}...")
        statistics_url = f"https://finance.yahoo.com/quote/{symbol_with_extension}/key-statistics"
        driver.get(statistics_url)
        market_cap = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[1]/div/div/div/div/table/tbody/tr[1]/td[2]').text
        print("Market Cap:", market_cap)
        share_price = driver.find_element(By.XPATH, '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]').text
        print("Share Price:", share_price)
        trailing_pe = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[1]/div/div/div/div/table/tbody/tr[3]/td[2]').text
        print("Trailing P/E Ratio:", trailing_pe)
        price_to_book = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[1]/div/div/div/div/table/tbody/tr[7]/td[2]').text
        print("Price/Book (mrq):", price_to_book)
        beta = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[2]/div/div[1]/div/div/table/tbody/tr[1]/td[2]').text
        print("Beta:", beta)
        fifty_two_week_high = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[2]/div/div[1]/div/div/table/tbody/tr[4]/td[2]').text
        print("52 Week High:", fifty_two_week_high)
        fifty_two_week_low = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[2]/div/div[1]/div/div/table/tbody/tr[5]/td[2]').text
        print("52 Week Low:", fifty_two_week_low)
        fifty_day_moving_avg = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[2]/div/div[1]/div/div/table/tbody/tr[6]/td[2]').text
        print("50-Day Moving Average:", fifty_day_moving_avg)
        enterprise_value = driver.find_element(By.XPATH, '//*[@id="Col1-0-KeyStatistics-Proxy"]/section/div[2]/div[1]/div/div/div/div/table/tbody/tr[2]/td[2]').text
        print("Enterprise Value:", enterprise_value)

        # Scraping sector from profile tab
        print(f"Scraping sector for symbol {symbol}...")
        profile_url = f"https://finance.yahoo.com/quote/{symbol_with_extension}/profile"
        driver.get(profile_url)
        sector = driver.find_element(By.XPATH, '//*[@id="Col1-0-Profile-Proxy"]/section/div[1]/div/div/p[2]/span[2]').text
        print("Sector:", sector)

        # Scraping full-time employees from profile page
        print(f"Scraping full-time employees for symbol {symbol}...")
        full_time_employees = driver.find_element(By.XPATH, '//*[@id="Col1-0-Profile-Proxy"]/section/div[1]/div/div/p[2]/span[6]').text
        print("Full Time Employees:", full_time_employees)

        return market_cap, share_price, trailing_pe, price_to_book, beta, fifty_two_week_high, fifty_two_week_low, fifty_day_moving_avg, enterprise_value, sector, full_time_employees
    except NoSuchElementException as e:
        print(f"Error: {e}")
        return None, None, None, None, None, None, None, None, None, None, None
    except Exception as e:
        print(f"Error scraping data for symbol {symbol}: {e}")
        return None, None, None, None, None, None, None, None, None, None, None

# Function to calculate the value for "indicator" field based on the given formula
def calculate_indicator(share_price, fifty_two_week_high, fifty_two_week_low):
    try:
        share_price_value = float(share_price.replace(',', ''))
        fifty_two_week_high_value = float(fifty_two_week_high.replace(',', ''))
        fifty_two_week_low_value = float(fifty_two_week_low.replace(',', ''))
        
        if share_price_value > (fifty_two_week_high_value - fifty_two_week_high_value * 0.05):
            return "Close to 52 week High"
        elif share_price_value < (fifty_two_week_low_value + fifty_two_week_low_value * 0.05):
            return "Close to 52 week low"
        else:
            return ""
    except Exception as e:
        print(f"Error calculating 'indicator' value: {e}")
        return ""

# Function to calculate the value for "indicator_2" field based on the given formula
def calculate_indicator_2(share_price, fifty_day_moving_avg):
    try:
        share_price_value = float(share_price.replace(',', ''))
        fifty_day_moving_avg_value = float(fifty_day_moving_avg.replace(',', ''))
        
        if share_price_value > fifty_day_moving_avg_value:
            return "Above 50 day moving avg"
        else:
            return "Below 50 day moving avg."
    except Exception as e:
        print(f"Error calculating 'indicator_2' value: {e}")
        return ""

# Main function to perform the tasks
def main():
    # Read the CSV file
    df = pd.read_csv("ind_nifty500list_usecase3.csv")
    
    # Initialize the WebDriver
    driver = initialize_driver()

    if driver:
        # Create a new DataFrame to store the results
        results_df = pd.DataFrame(columns=['Company Name', 'Industry', 'Sector', 'Ticker', 'Share Price', 'Market Cap', 'Enterprise Value', 'Trailing P/E', 'PB', 'Beta', '52 Week High', '52 Week Low', '50-Day Moving Average', 'No. of employees', 'Indicator', 'Indicator_2'])

        # Iterate through each symbol in the DataFrame
        for index, row in df.iterrows():
            symbol = row['Symbol']
            company_name = row['Company Name']
            industry = row['Industry']
            print(f"Processing symbol: {symbol}...")
            
            # Scrape data for the symbol
            market_cap, share_price, trailing_pe, price_to_book, beta, fifty_two_week_high, fifty_two_week_low, fifty_day_moving_avg, enterprise_value, sector, full_time_employees = scrape_company_data(driver, symbol)
            if market_cap is not None and share_price is not None and trailing_pe is not None and price_to_book is not None and beta is not None and fifty_two_week_high is not None and fifty_two_week_low is not None and fifty_day_moving_avg is not None and enterprise_value is not None and sector is not None and full_time_employees is not None:
                # Calculate additional fields
                indicator = calculate_indicator(share_price, fifty_two_week_high, fifty_two_week_low)
                indicator_2 = calculate_indicator_2(share_price, fifty_day_moving_avg)
                
                # Append the data to the results DataFrame
                results_df.loc[len(results_df)] = [company_name, industry, sector, symbol, share_price, market_cap, enterprise_value, trailing_pe, price_to_book, beta, fifty_two_week_high, fifty_two_week_low, fifty_day_moving_avg, full_time_employees, indicator, indicator_2]

        # Save the results to a new CSV file
        results_df.to_csv("company_data_29_Mar.csv", index=False)
        print("Results saved to company_data.csv")
        
        # Quit the WebDriver
        driver.quit()
        print("WebDriver closed.")
    else:
        print("WebDriver initialization failed. Exiting.")

# Execute the main function
if __name__ == "__main__":
    main()