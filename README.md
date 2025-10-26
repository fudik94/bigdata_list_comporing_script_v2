# Excel File Comparator

This project compares two large Excel files by company code.  
It finds which companies exist in each file, detects name mismatches, and creates a clean merged report.



## Features

- Reads two Excel files
- Cleans and normalizes company codes and names
- Detects which entries exist in each file
- Flags different names for the same code
- Handles hundreds of thousands of rows efficiently
- Saves a final report in Excel format



## Requirements

- Python 3.9 or higher  
- pandas library  

Install dependencies:

```bash
pip install pandas openpyxl
