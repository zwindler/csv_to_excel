# csv_to_excel
A python tool to merge multiple csv files into one xls file with multiple sheets

# Author
The script has initially been writen by Sujit Pal and can be found on http://sujitpal.blogspot.fr/2007/02/python-script-to-convert-csv-files-to.html
Sujit version doesn't allow multiple csv file as input, but the python module he used does. I had very little code to change so my work here is minimal. Please thank him.

# Usage

```python
python csv_to_excel.py file1.csv #script works just like before with one csv file
python csv_to_excel file1.csv file2.csv file3.csv #but here you can use this syntax to merge multiple .csv (one sheet for each .csv)
python csv_to_excel file*.csv #wildcard also works
```
