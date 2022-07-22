# CLISC<sup>3</sup> - WORK IN PROGRESS
**Command Line Interface Spreadsheet Count Convert & Compare (& Archive)**

A small Windows Exe program made in C#. It is as a hobby project. It might have use cases in digital archiving of spreadsheets.

## How to use
Download the executable here. In the terminal change directory to the folder where CLISC.exe is. Then, to execute the program input:

```
.\CLISC.exe [YourArguments]
```

Replace [YourArguments] with one of the following*:

```
Count 'path to input dir' 'path to output dir' Recurse='Yes'/'No'
```
```
Count&Convert 'path to input dir' 'path to output dir' Recurse='Yes'/'No'
```
```
Count&Convert&Compare 'path to input dir' 'path to output dir' Recurse='Yes'/'No'
```
```
Count&Convert&Compare&Archive 'path to input dir' 'path to output dir' Recurse='Yes'/'No'
```
<sub>*Remove '...' around arguments</sub>

<sub>*You must input arguments in the order above, and you cannot leave out an argument</sub>

## Program behavior

### General
* Bulk convert spreadsheets in a directory
* Include or exclude subdirectories recursively
* Output results in CSV logs
* Password protected or otherwise unreadable spreadsheets will not be converted or compared, but they will be copied to the new archivable directory

### Count
* Count number of spreadsheets in directory by file format 
  - Accepted file extensions: .fods, .ods, .ots, .xla, .xlam, .xls, .xlsb, .xlsm, .xlsx, .xlt, .xltm, .xltx
  - Office Open XML file formats of Transitional and Strict conformance can be identified

### Convert
Convert any spreadsheet to .xlsx (Excel, Office Open XML Transitional conformance)

* The following file formats are supported
  - Office Open XML with extensions .xlsb, .xlsm, .xltm, .xltx and .xlsx with Strict conformance*
  - Legacy Microsoft Excel with extensions .xls and .xlt* (feature not working)
  - OpenDocument with extensions .fods, .ods and .ots
* Output all conversions in subdirectories named n+1
* Rename all conversions n+1.xlsx

<sub>*.xla and .xlam file extensions are Microsoft Excel Add-in files and cannot contain worksheet cell information. Therefore, they are excluded from conversion but will be copied to the new archive directory, if arhciving is selected</sub>

<sub>*.xlsx with Transitional conformance will only be converted if arhciving is selected</sub>

### Compare
Compare original and converted spreadsheets to log differences* in

* Workbook cell values
* File size

<sub>*The program can currently not compare cell formatting, embedded objects, charts and other advanced spreadsheet features.</sub>

### Advanced archival requirements (use argument Archive=Yes)
The program supports the conversion of spreadsheets to meet a data quality level, that will enable you to open your spreadsheets many years from now. Enabling advanced archival requirements will

* Package spreadsheets and metadata in a new archive directory
* Include copies of the original spreadsheets
* Validate spreadsheets against their file format standards (Office Open XML and OpenDocument) (feature not working)
* Remove formula linking cells to other local spreadsheets but keep the cell values (feature not working)
* Remove external data connections but keep snapshot of data (feature not working)
* Remove RealTimeData (RTD) functions but keep snapshot of data (feature not working)
* Alert if spreadsheet has embedded objects (feature not working)
* Calculate file checksums
* Zip the archive directory

## Dependencies
Prerequisite software for the program to work with these functions.

### Convert
* [LibreOffice](https://www.libreoffice.org/)
  - If you want to convert OpenDocument spreadsheets
  - You need to install program in its default directory
  - The program is free
* [EPPLUS6](https://www.epplussoftware.com/)
  - If you want to convert legacy Excel file formats
  - You do NOT need to install
  - You need to purchase license

### Compare
* [Beyond Compare 4](https://www.scootersoftware.com/)
  - If you want to use the compare function
  - You need to install program in its default directory
  - You need to purchase license
