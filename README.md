# CLISC
**Command Line Interface Spreadsheet Count Convert & Compare (& Archive)**

A small Windows console application made in C#. It is a hobby project. The app might have use cases in digital archiving of spreadsheets.

:rainbow_flag: **General**
* Bulk convert spreadsheets in a directory to .xlsx (Transitional conformance)
* Include or exclude subdirectories recursively
* Output results in a new directory with logs in CSV

:heavy_plus_sign: **Count**

Count number of spreadsheets in directory by file format. 
* Accepted file extensions: .fods, .ods, .ots, .xla, .xlam, .xls, .xlsb, .xlsm, .xlsx, .xlt, .xltm, .xltx
* Office Open XML file formats of Transitional and Strict conformance can be counted separately

:magic_wand: **Convert**

Convert any spreadsheet[^1] to .xlsx (Transitional conformance).
* Office Open XML (Excel) with extensions .xlsb, .xlsm, .xltm, .xltx and .xlsx with Strict conformance[^2]
* Legacy Microsoft Excel with extensions .xls and .xlt
* OpenDocument with extensions .fods, .ods and .ots

:mag: **Compare**

Compare original and converted spreadsheets to log differences.[^3]
* Workbook cell values
* File size

:file_cabinet: **Archive**

The program can convert, package and describe spreadsheets to meet a data quality level, that will enable you to open your spreadsheets many years from now. 
* Convert any spreadsheet[^1] to both .xlsx (Transitional conformance) and .ods
* Package spreadsheets and metadata in a new archive directory
* Output all conversions in subdirectories named n+1
* Rename all conversions n+1.xlsx and n+1.ods
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

## How to use
Download the executable [here](https://github.com/Asbjoedt/CLISC/releases). In your terminal change directory to the folder where CLISC.exe is. Then, to execute the program input:

```
.\CLISC.exe [your_arguments]
```

Create your arguments from the following list:

**Functions to use** (required, pick one of the four)
```
--function count
--function count&convert
--function count&convert&compare
--function count&convert&compare&archive
```
**Input directory** (required)
```
--inputdir ["path to input directory"]
```
**Output directory** (required)
```
--outputdir ["path to output directory"]
```
**Include subdirectories from input directory** (optional, by default false)
```
--recurse true
```
**Example of full usage**
```
.\CLISC.exe --function count&convert&compare&archive --inputdir "c:\my_path" --outputdir "c:\my_path" --recurse true
```
**or shorter**
```
.\CLISC.exe -f count&convert&compare&archive -i "c:\my_path" -o "c:\my_path" -r true
```

If you want to test the application, a sample dataset is provided [here](https://github.com/Asbjoedt/CLISC/tree/master/Test_Data).

## Libraries and software used
* [Beyond Compare 4](https://www.scootersoftware.com/index.php), Copyright (c) 2022 Scooter Software, Inc.
* [CommandLineParser](https://github.com/commandlineparser/commandline), MIT License, Copyright (c) 2005 - 2015 Giacomo Stelluti Scala & Contributor
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [NPOI](https://github.com/nissl-lab/npoi), Apache License 2.0, no changes
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation

[^1]: File extensions .xla and .xlam are Microsoft Excel Add-in files and cannot contain worksheet cell information. Therefore, they are excluded from conversion but will be copied to the new archive directory, if archiving is selected.
[^2]: Conversion to file extension .xlsx with Strict conformance are currently not supported.
[^3]: The program can currently not compare cell formatting, embedded objects, charts and other advanced spreadsheet features.
