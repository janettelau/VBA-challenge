# VBA-challenge
Using VBA scripting to analyze generated stock market data

## Description
The script loops through each row for all stocks for each quarter (each worksheet) to calculate the
- Quarterly Change: The opening price at the beginning of the quarter is subtracted from the closing price at the end of the quarter.
- Percent Change: The quarterly change divided by the opening price at the beginning of the quarter, in %.
- Total Stock Volume: The sum of daily volumes for each stock throughout the quarter.
It then outputs the results in a summary table on the same worksheet.

The script also loops through each stock to identify the
- Greatest % Increase and Greatest % Decrease from the Percent Change values.
- Greatest Total Volume from the Total Stock Volume values.
Results are printed on the same worksheet as well.

The final outputs are the screenshot images Q1, Q2, Q3, and Q4.
