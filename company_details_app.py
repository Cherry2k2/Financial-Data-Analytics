import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Load your first database (Excel sheet) containing basic company details
def load_basic_data():
    return pd.read_csv('company_data.csv')

# Load your second database (Excel sheet) containing historical data
def load_historical_data():
    return pd.read_excel('360ONE.NS.xlsx')

# Main function to search for a company and display details from both sheets
def main():
    st.title('Company Details Search')

    # Load basic and historical data
    basic_data = load_basic_data()
    historical_data = load_historical_data()

    # User input for company name
    company_name = st.text_input('Enter company name:')

    # Search for company and display details
    if st.button('Search'):
        if company_name:
            # Display details from the first sheet for the specific company
            company_basic_details = basic_data[basic_data['Company Name'].str.contains(company_name, case=False)]
            if not company_basic_details.empty:
                st.write("**Company Details Today:**")
                st.write(company_basic_details)
            else:
                st.write('Company details not found in the first file.')

            # Display details from the second sheet
            company_historical_details = historical_data[historical_data['Company Name'].str.contains(company_name, case=False)]
            if not company_historical_details.empty:
                st.write("**Company Details for analysis:**")
                st.write(company_historical_details)

                # Visualize historical data for the selected company from the second sheet
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.plot(company_historical_details['Date'], company_historical_details['Open'], label='Open')
                ax.plot(company_historical_details['Date'], company_historical_details['Close'], label='Close')
                ax.plot(company_historical_details['Date'], company_historical_details['High'], label='High')
                ax.plot(company_historical_details['Date'], company_historical_details['Low'], label='Low')
                ax.set_xlabel('Date')
                ax.set_ylabel('Price')
                ax.set_title('Historical Prices of Companies')
                ax.legend()
                ax.tick_params(axis='x', rotation=45)
                st.pyplot(fig)

                # Print insights
                st.write("**Insights:**")
                st.write("- The graph shows the historical prices (Open, Close, High, Low) of the selected company over time.")
                st.write("- You can observe the fluctuations and trends in the stock prices.")
            else:
                st.write('Company not found in the second file.')
        else:
            st.write('Please enter a company name.')

if __name__ == '__main__':
    main()