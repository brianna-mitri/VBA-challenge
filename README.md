# VBA Challenge
## Description
Analyze quarterly stock data (ticker, date, prices, and volume) across Excel sheets using VBA to reveal the following:
  - Each stock's quarterly price changes/percentages and total stock volume for the quarter.
  - Stocks with the greatest quarterly price percentage increase/decrease and highest stock volume.
    
Each Excel sheet represents one quarter of data; however, the first sheet contains the instructions along with the two buttons to run each macro.

## Run Macros
Run the macros by pressing the following buttons found on the ***"Instructions"*** sheet in the Excel file:
  - **Re-sort Data** _(macro: resort_data)_: sort the data both alphabetically and chronologically to ensure output tables are computed correctly. Useful when adding new data.
  - **Get Output Tables** _(macro: stock_output_tables)_: generate two tables summarizing and highlighting stock data for each quarter. Both table outputs are found on each sheet.

## Files
  - **BM_Multiple_Year_Stock_Data.xlsm:** includes quarterly stock data and macros to sort the data and generate summary tables.
  - **BM_VBA Challenge_Results Per Quarter.pdf:** includes screenshots of the output tables per quarter.

> [!NOTE]
The generation of output tables (_stock_output_tables_ macro) relies on the quarterly stock data being organized both alphabetically and chronologically. To ensure accurate output tables, please use the _resort_data_ macro to re-sort the data after adding new data or reorganizing the data.
