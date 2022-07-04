# CLISC
**Command Line Interface Spreadsheet Count Convert & Compare**

A small Windows Exe program made in C#. It is as a hobby project. It might have use cases in digital archiving of spreadsheets.

## Program behavior
**Count**
* Count number of spreadsheets in directory by file format (file extension)

**Convert**
* Convert XLS, XLT, XLAM, XLSB, XLTX, XLSM, XLTM to XLSX (OOXML Transitional conformance)
* Output all conversions in a new directory with new subdirectories named n+1
* Rename all conversions n+1.xlsx

**Compare**
* Compare the spreadsheets to log workbook and checksum differences between input and output file formats

**General**
* Include or exclude subdirectories recursively
* Output results in CSV logs

## How to use
In the terminal navigate (cd) to directory of CLISC.exe. To execute the program, input:

```
.\CLISC.exe [YourArgument]
```

Replace [YourArgument] with one of the following arguments:

```
Count 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
```
Count&Convert 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
```
Count&Convert&Compare 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```

## Dependencies
Prerequisite tools for the program to work.

**Compare**
* Microsoft Spreadsheet Compare, which is included in Microsoft Office Professional Plus 2013, 2016, 2019
