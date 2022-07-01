# CLISC
## Program behavior:
* Count Excel spreadsheets in directory by file format
* Convert XLS, XLT, XLAM, XLSB, XLTX, XLSM, XLTM to XLSX (OOXML Transitional conformance)
* Output all conversions in a new directory with new subdirectories named n+1
* Rename all conversions n+1.xlsx
* Compare the results to log workbook and checksum differences between input and output file formats
* Output log in CSV

## Available arguments:
```
Count 'Filepath to directory' -Recursive
```
```
Count&Convert 'Filepath to directory' -Recursive
```
```
Count&Convert&Compare 'Filepath to directory' -Recursive
```
