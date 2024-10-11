# VBA Challenge
## Description
Analyze quarterly stock data (ticker, date, prices, and volume) across Excel sheets using VBA to reveal the following:
  - Each stock's quarterly price changes/percentages and total stock volume for the quarter.
  - Stocks with the greatest quarterly price percentage increase/decrease and highest stock volume.
    
Each Excel sheet represents one quarter of data; however, an optional "Instructions" sheet would be skipped over.

## Macros
Here is what each of the macros does:
  - **Re-sort Data** _(macro: resort_data)_: sort the data both alphabetically and chronologically to ensure output tables are computed correctly. Useful when adding new data.
  - **Get Output Tables** _(macro: stock_output_tables)_: generate two tables summarizing and highlighting stock data for each quarter. Both table outputs are found on each sheet.

## Expected Data Format
The quarterly stock data on each Excel sheet is expected to be organized from columns A to G in the following order: Ticker, Date, Opening Price, Highest Price, Lowest Price, Closing Price, and Volume[^1].

The macros account for an Excel sheet, "Instructions", to be skipped if present.

[^1]: Raw data for the Highest Price and Lowest Price columns are not used in the analysis so (if using the macros as is) those could be replaced with different values.

## Files
  - **BM_Macro_ResortData.bas:** includes the VBA script for the re-sort data macro.
  - **BM_Macro_OutputTables.bas:** includes the VBA script for the get output tables macro.
  - **BM_Multiple_Year_Stock_Data.xlsm:** includes generated quarterly stock data and an "Instructions" sheet with buttons to run each macro.
  - **BM_VBA Challenge_Results Per Quarter.pdf:** includes screenshots of the macros' results per quarter.

> [!NOTE]
The generation of output tables (_stock_output_tables_ macro) relies on the quarterly stock data being organized both alphabetically and chronologically. To ensure accurate output tables, please use the _resort_data_ macro to re-sort the data after adding new data or reorganizing the data.
