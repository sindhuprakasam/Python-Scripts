Python Scripts
=============
Data Engineering, Report Creations, Automation scripts

### Report Automation
- report generation
- report for TM

##### Report Generation:
This Python script is to read data from a csv file and to create excel dashboard with all the metrics/summary details for the users.

Input: input_report_generation.csv

Output: report_generation.xlsx

#####Description:
The script reads the input csv file and all the previous days data saved in pickle files. Goal is to create a mini dashboard with the data available in these files. For that i need to create few additional columns linking the current day and the previous day's data. All the calculations for the data have been done with Pandas. Created a summary page with graphs/charts to get a overall understanding of the data given using Python's matplotlib.Fianlly saving the updated data and current data as pickle files for future use.

