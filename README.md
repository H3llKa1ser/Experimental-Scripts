# Experimental Scripts

## Required dependencies:

- openpyxl

### Install dependency

    sudo apt install python3-openpyxl

## Example:

### 1) Parse data from .txt to .csv file

    python3 txt2csv_parser.py -i file.txt -o csvparsed.csv

### 2) Parse data from the parsed .csv to .xlsx workbook file

    python3 csv2xlsx_parser.py -i csvparsed.csv -o workbook.xlsx

### 3) "Beautify" .xlsx workbook file, making it easier to sort and edit data within it

    python3 beautexcel.py workbook.xlsx
