# Excel-Automation-using-Python-mini-scripts-
#### Short python scripts to automate excel sheets just by changing the file name, column name, etc
## Creating separate workbooks from one using values of a particular column
> Dataset: sales_data.xlsx

Using values in a particular column, n number of different excel sheets will be created which will contain all it's respective data

n : number of different values in that column

Here in sales_data, differnt sheets are created based on product category column

## Merging various sheets into one having similar data on one column
>Dataset: shift-data, third-shift-data

Using a particular column , different sheets will be merged into one
(Consider this just as opposite of the one above)

Here data is separated into various sheets according to various shifts, this will merge it into one and then we can analyse it using one single dataframe
