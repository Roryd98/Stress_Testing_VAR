#%% Import Library

import pandas as pd
from datetime import datetime
from IPython.display import display
import numpy as np
import os

#%%
#%% Load and Set Data

# Load the excel file - set file path as per actual path.
file_path = "/Users/rorymcmullan/Downloads/Project_Data_2108.xlsm"


# load both sheets into the dataframe

prices_df = pd.read_excel(file_path, sheet_name='Data')
returns_df  =pd.read_excel(file_path, sheet_name='Returns')

# Convert 'Date' column to datetime and set as index
prices_df['Date'] = pd.to_datetime(prices_df['Date'], format='%d/%m/%Y')
returns_df['Date'] = pd.to_datetime(returns_df['Date'], format='%d/%m/%Y')

prices_df.set_index('Date', inplace=True)
returns_df.set_index('Date', inplace=True)

#Display the df

#display(prices_df.head())
#display(returns_df.head())

#%%
# Worst 1 day return for multiple periods
"""FIND THE WORST 1 DAY RETURN FOR MULTIPLE CUSTOM PERIODS"""

# Ensure the index is sorted
returns_df.sort_index(inplace=True)

# Define multiple periods with start and end dates. periods have been defined in this block only as represents all dates throughout the script as seen here.  

periods = [
    ('17/03/2022', '23/07/2024'),
    ('24/02/2020', '24/02/2021'),  
    ('23/06/2016', '23/06/2017'),  
    ('22/05/2013', '30/01/2015'),  
    ('16/04/2010', '07/05/2010'), 
    ('09/09/2009', '13/10/2009'),
    ('08/09/2009', '18/09/2009'),  
    ('09/02/2009', '09/03/2009'),
    ('20/07/2007', '20/08/2007'),
    ('17/06/2003', '13/08/2003'),
    ('10/09/2001', '21/09/2001')
    ]
"""Note the periods above only have to be defined once and are saved by the variable 'periods'. To adjust periods,
change these dates or else assign a new variable for example periods_new, and enter new date ranges."""

# Convert the date strings to datetime objects with the correct format
periods = [(pd.to_datetime(start, format='%d/%m/%Y'), pd.to_datetime(end, format='%d/%m/%Y')) for start, end in periods]

# Iterate over each period to find the worst 1-day return
for period in periods:
    start_date, end_date = period  # Unpack the tuple into start_date and end_date

    # Filter the returns for the given dates
    filtered_returns_df = returns_df.loc[start_date:end_date]

    # Find the worst 1-day return for each column (each index)
    Minimum_Daily_Returns = filtered_returns_df.min()

    # Display the results
    print(f"Worst 1-Day Returns from {start_date} to {end_date}:")
    print(Minimum_Daily_Returns)
#%%
# Maximum Drawdown

prices_df.sort_index(inplace=True)

# Convert the date strings to datetime objects with the correct format
periods = [(pd.to_datetime(start, format='%d/%m/%Y'), pd.to_datetime(end, format='%d/%m/%Y')) for start, end in periods]

# Function to calculate the maximum drawdown
def calculate_max_drawdown(prices):
    try:
        # Calculate the cumulative returns
        cumulative = prices / prices.iloc[0]

        # Calculate the running maximum
        running_max = cumulative.cummax()

        # Calculate the drawdown
        drawdown = (cumulative - running_max) / running_max

        # Find the maximum drawdown
        max_drawdown = drawdown.min()

        return max_drawdown
    except Exception as e:
        print(f"Error calculating max drawdown for {prices.name}: {e}")
        return None

# Loop through each period to calculate the maximum drawdown
for start_date, end_date in periods:
    # Filter the prices for the given dates
    filtered_prices_df = prices_df.loc[start_date:end_date]
    
    # Apply the function to each column
    Maximum_Drawdowns = filtered_prices_df.apply(calculate_max_drawdown)
    
    # Print the maximum drawdowns for the current period
    print(f"Maximum Drawdowns from {start_date} to {end_date}:")
    print(Maximum_Drawdowns)
  
#%%
# Full Return
"""FIND THE FULL RETURN OF AN INDEX WITHIN A CUSTOM PERIOD"""

# Drop rows where the index contains NaT
prices_df = prices_df[prices_df.index.notna()]

# Ensure the index is sorted and remove duplicates
prices_df = prices_df[~prices_df.index.duplicated(keep='first')]
prices_df.sort_index(inplace=True)

# Check if the index is monotonic after cleaning
if prices_df.index.is_monotonic_increasing:
    print("Index is now monotonic increasing.")
else:
    print("Index is still not monotonic increasing after cleaning.")


# Sort the DataFrame by the index to ensure proper slicing
prices_df.sort_index(inplace=True)


# Convert the date strings to datetime objects with the correct format
periods = [(pd.to_datetime(start, format='%d/%m/%Y'), pd.to_datetime(end, format='%d/%m/%Y')) for start, end in periods]

# Loop through each period to calculate the full period returns
for start_date, end_date in periods:
    # Convert start_date and end_date to datetime
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    
    # Ensure that we find the nearest available dates if the exact dates don't exist
    nearest_start_date = prices_df.index.asof(start_date)
    nearest_end_date = prices_df.index.asof(end_date)

    # Handle cases where no valid dates are found
    if pd.isna(nearest_start_date) or pd.isna(nearest_end_date):
        print(f"No data available for the period {start_date.date()} to {end_date.date()}. Skipping this period.")
        continue

    # Filter the prices for the given nearest dates
    filtered_prices_df = prices_df.loc[nearest_start_date:nearest_end_date]

    # Get the prices at the start and end of the period
    start_prices = filtered_prices_df.loc[nearest_start_date]
    end_prices = filtered_prices_df.loc[nearest_end_date]

    # Calculate the return
    Full_Period_Return = (end_prices - start_prices) / start_prices

    # Print the full period return for the current period
    print(f"Full Period Returns from {nearest_start_date.date()} to {nearest_end_date.date()}:")
    print(Full_Period_Return)
    print()  # Adding a blank line for readability
  
#%%
# Define multiple periods with start and end dates

# Convert the date strings to datetime objects with the correct format
periods = [(pd.to_datetime(start, format='%d/%m/%Y'), pd.to_datetime(end, format='%d/%m/%Y')) for start, end in periods]

# Initialize a list to store the combined results
combined_results = []

# Function to calculate the maximum drawdown
def calculate_max_drawdown(prices):
    cumulative = prices / prices.iloc[0]
    running_max = cumulative.cummax()
    drawdown = (cumulative - running_max) / running_max
    return drawdown.min()

# Loop through each period to calculate all metrics
for start_date, end_date in periods:
    # Filter the data for the given dates
    filtered_prices_df = prices_df.loc[start_date:end_date]
    filtered_returns_df = returns_df.loc[start_date:end_date]
    
    # Calculate Worst Daily Return
    min_daily_return = filtered_returns_df.min()

    # Calculate Max Drawdown
    max_drawdown = filtered_prices_df.apply(calculate_max_drawdown)
    
    # Calculate Full Period Return
    start_prices = filtered_prices_df.loc[start_date]
    end_prices = filtered_prices_df.loc[end_date]
    full_period_return = (end_prices - start_prices) / start_prices
    
    # Prepare the rows for each metric
    date_range_str = f"{start_date} - {end_date}"
    combined_results.append([date_range_str, 'Worst Daily Return (%)'] + min_daily_return.tolist())
    combined_results.append([date_range_str, 'Max Drawdown (%)'] + max_drawdown.tolist())
    combined_results.append([date_range_str, 'Full Period Return (%)'] + full_period_return.tolist())

# Create a DataFrame from the results
columns = ['Date Range', 'Metric'] + min_daily_return.index.tolist()
final_results_df = pd.DataFrame(combined_results, columns=columns)

# Output file name
output_file = 'financial_metrics_Final.csv'

# Write data to CSV file
final_results_df.to_csv(output_file, index=False)

print(f"Results have been saved to {output_file}")
