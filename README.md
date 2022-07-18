# CLISC<sup>3</sup> - WORK IN PROGRESS
**Command Line Interface Spreadsheet Count Convert & Compare**

A small Windows Exe program made in C#. It is as a hobby project. It might have use cases in digital archiving of spreadsheets.

## Program behavior
The program has three main functions.

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

### General
* Bulk convert spreadsheets in a directory
* Include or exclude subdirectories recursively
* Output results in CSV logs
* Password protected or otherwise unreadable spreadsheets will not be converted or compared, but they will be copied to the new archivable directory

## Dependencies
Prerequisite software for the program to work with the following functions.

### Convert
* LibreOffice (free, you need to install)
* EPPLUS6 (you need to purchase license, you do NOT need to install)

### Compare
* Beyond Compare 4 (you need to purchase license, you need to install)

## How to use
In the terminal change directory to the folder where CLISC.exe is. Then, to execute the program, input:

```
.\CLISC.exe [YourArgument]
```

Replace [YourArgument] with one of the following arguments*:

```
Count 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
```
Count&Convert 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
```
Count&Convert&Compare 'Filepath to input directory' 'Filepath to output directory' Recursive=Yes/No
```
<sub>*Input and output directorices are not allowed to be identical.</sub>

<sub>*Delete ' ' around input and output directories.</sub>
