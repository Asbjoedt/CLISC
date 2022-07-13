# CLISC - WORK IN PROGRESS
**Command Line Interface Spreadsheet Count Convert & Compare**

A small Windows Exe program made in C#. It is as a hobby project. It might have use cases in digital archiving of spreadsheets.

## Program behavior
Three functions.

### Count
* Count number of spreadsheets in directory by file format (file extension)

### Convert
Convert any spreadsheet to XLSX (OfficeOpen XML Transitional conformance)

* The following file formats are supported
  - OfficeOpen XML with extensions XLSB, XLTX, XLSM and XLTM<sup>*</sup>
  - Legacy Microsoft Excel with extensions XLS and XLT<sup>*</sup> (feature not working)
  - OpenDocument with extensions FODS, ODS and OTS (feature not working)
* Output all conversions in subdirectories named n+1
* Rename all conversions n+1.xlsx
* Your spreadsheets are packaged in a new archivable directory, which includes copies of the original spreadsheets

<sup>*XLA and XLAM file extensions are Microsoft Excel Add-in files and cannot contain worksheet cell information. Therefore, they are excluded from conversion but will be copied to the new archivable directory.</sup>

### Compare
* Compare the spreadsheets to log workbook, file size and checksum differences between input and output file formats

### General
* Include or exclude subdirectories recursively
* Output results in CSV logs
* Password protected or otherwise unreadable spreadsheets will not be converted or compared, but they will be copied to the new archivable directory

## Dependencies
Prerequisite software for the program to work with all functions.

**Compare**
* Microsoft Spreadsheet Compare, which is included in Microsoft Office Professional Plus 2013, 2016, 2019

## How to use
In the terminal navigate (cd) to directory of CLISC.exe. To execute the program, input:

```
.\CLISC.exe [YourArgument]
```

Replace [YourArgument] with one of the following arguments:

*-- Input and output directorices are not allowed to be identical.*

*-- Delete ' ' around input and output directories.*

```
Count 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
```
Count&Convert 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
```
Count&Convert&Compare 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
