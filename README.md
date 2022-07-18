# CLISC<sup>3</sup> - WORK IN PROGRESS
**Command Line Interface Spreadsheet Count Convert & Compare**

A small Windows Exe program made in C#. It is as a hobby project. It might have use cases in digital archiving of spreadsheets.

## How to use
Download the executable here. In the terminal change directory to the folder where CLISC.exe is. Then, to execute the program input:

```
.\CLISC.exe [YourArgument]
```

Replace [YourArgument] with one of the following arguments*:

```
Count 'path to input directory' 'path to output directory' Recursive=Yes/No
```
```
Count&Convert 'path to input directory' 'path to output directory' Recursive=Yes/No
```
```
Count&Convert&Compare 'path to input directory' 'path to output directory' Recursive=Yes/No
```
<sub>*Input and output directorices are not allowed to be identical.</sub>

<sub>*Delete ' ' around input and output directories.</sub>

## Program behavior

### General
* Bulk convert spreadsheets in a directory
* Include or exclude subdirectories recursively
* Output results in CSV logs
* Password protected or otherwise unreadable spreadsheets will not be converted or compared, but they will be copied to the new archivable directory

### Count
* Count number of spreadsheets in directory by file format (file extension)

### Convert
Convert any spreadsheet to XLSX (Excel, Office Open XML Transitional conformance)

* The following file formats are supported
  - Office Open XML with extensions XLSB, XLTX, XLSM and XLTM*
  - Legacy Microsoft Excel with extensions XLS and XLT* (feature not working)
  - OpenDocument with extensions FODS, ODS and OTS
* Output all conversions in subdirectories named n+1
* Rename all conversions n+1.xlsx
* Your spreadsheets are packaged in a new archivable directory, which includes copies of the original spreadsheets

<sub>*XLA and XLAM file extensions are Microsoft Excel Add-in files and cannot contain worksheet cell information. Therefore, they are excluded from conversion but will be copied to the new archivable directory.</sub>

### Compare
Compare original and converted spreadsheets to log differences* in

* Workbook cell values
* File size
* File checksum

<sub>*The program can currently not compare cell formatting, embedded objects, charts and other advanced spreadsheet features.</sub>

### Advanced archival requirements (feature not working)

* Validate spreadsheet against its file format standard (Office Open XML and OpenDocument)
* Remove formula linking celss to other local spreadsheets but keep the cell values
* Remove external data connections but keep snapshot of data
* Remove RealTimeData (RTD) functions but keep snapshot of data
* Alert if spreadsheet has embedded objects

## Dependencies
Prerequisite software for the program to work with these functions.

### Convert
* LibreOffice
  - If you want to convert OpenDocument spreadsheets in file extensions FODS, ODS and OTS
  - You need to install program in its default directory
  - The program is free
* EPPLUS6
  - If you want to convert legacy Excel file formats in extensions XLS and XLT
  - You do NOT need to install
  - You need to purchase license

### Compare
* Beyond Compare 4 
  - If you want to use the compare function
  - You need to install program in its default directory
  - You need to purchase license
