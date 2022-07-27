# CLISC
**Command Line Interface Spreadsheet Count Convert & Compare (& Archive)**

A small Windows console application made in C#. It is as a hobby project. The app might have use cases in digital archiving of spreadsheets.

## How to use
Download the executable [here](https://github.com/Asbjoedt/CLISC/releases). In the terminal change directory to the folder where CLISC.exe is. Then, to execute the program input:

```
.\CLISC.exe [YourArguments]
```

Replace [YourArguments] with one of the following[^1]:

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

If you wish to test the application, a sample dataset is provided [here](https://github.com/Asbjoedt/CLISC/tree/master/Test_Data).

## Program behavior

:rainbow_flag: **General**
* Bulk convert spreadsheets in a directory to .xlsx Transitional conformance
* Include or exclude subdirectories recursively
* Output results in a new directory with logs in CSV

:heavy_plus_sign: **Count**

Count number of spreadsheets in directory by file format. 
* Accepted file extensions: .fods, .ods, .ots, .xla, .xlam, .xls, .xlsb, .xlsm, .xlsx, .xlt, .xltm, .xltx
* Office Open XML file formats of Transitional and Strict conformance can be counted separately

:magic_wand: **Convert**

Convert any spreadsheet[^2] to .xlsx (Office Open XML Transitional conformance).
* Office Open XML with extensions .xlsb, .xlsm, .xltm, .xltx and .xlsx with Strict conformance[^3]
* Legacy Microsoft Excel with extensions .xls and .xlt
* OpenDocument with extensions .fods, .ods and .ots

:mag: **Compare**

Compare original and converted spreadsheets to log differences.[^4]
* Workbook cell values
* File size

:file_cabinet: **Archive**

The program can convert, package and describe spreadsheets to meet a data quality level, that will enable you to open your spreadsheets many years from now. 
* Package spreadsheets and metadata in a new archive directory
* Output all conversions in subdirectories named n+1
* Rename all conversions n+1.xlsx
* Include copies of the original spreadsheets, this include password protected or otherwise unreadable spreadsheets
* Validate spreadsheet against its file format standard (Office Open XML)
* Remove formula linking cells to other local spreadsheets but keep snapshot of cell values (feature not working)
* Remove external data connections but keep snapshot of cell values (feature not working)
* Remove RealTimeData (RTD) functions but keep snapshot of cell values (feature not working)
* Alert if spreadsheet has embedded objects (feature not working)
* Calculate file checksums
* Zip the archive directory

## Dependencies
Prerequisite software for the program to work with these functions.

:warning: **Convert**
* [LibreOffice](https://www.libreoffice.org/)
  - If you want to convert OpenDocument spreadsheets
  - You need to install program in its default directory
  - The program is free

:warning: **Compare**
* [Beyond Compare 4](https://www.scootersoftware.com/)
  - If you want to use the compare function
  - You need to install program in its default directory
  - You need to purchase license

[^1]: Remove '...' around arguments. You must input arguments in the order above, and you cannot leave out an argument.
[^2]: File extensions .xla and .xlam are Microsoft Excel Add-in files and cannot contain worksheet cell information. Therefore, they are excluded from conversion but will be copied to the new archive directory, if arhciving is selected.
[^3]: File extension .xlsx with Transitional conformance will only be converted if archiving is selected.
[^4]: The program can currently not compare cell formatting, embedded objects, charts and other advanced spreadsheet features.
