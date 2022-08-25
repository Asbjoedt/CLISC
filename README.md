# CLISC
**Command Line Interface Spreadsheet Count, Convert, Compare & Archive**

A small Windows console application made in C#. It is a hobby project. The app might have use cases in digital archiving of spreadsheets. For more information, see the [wiki](https://github.com/Asbjoedt/CLISC/wiki). For graphical user interface, see child repository [GUISC](https://github.com/Asbjoedt/GUISC).

:rainbow_flag: **General**

* Batch convert spreadsheets in a directory to .xlsx
* Include or exclude subdirectories recursively
* Output results in a new directory with logs in CSV

:heavy_plus_sign: **Count**

Count number of spreadsheets in directory by file format. 
* Accepted file extensions: .gsheet, .fods, .numbers, .ods, .ots, .xla, .xlam, .xls, .xlsb, .xlsm, .xlsx, .xlt, .xltm, .xltx
* .xlsx of Transitional and Strict conformance can be counted separately

:magic_wand: **Convert**

Convert any spreadsheet[^1] to .xlsx (Transitional conformance).
* Office Open XML (Excel) with extensions .xlsb, .xlsm, .xltm, .xltx and .xlsx with Strict conformance[^2]
* Legacy Microsoft Excel with extensions .xls and .xlt
* OpenDocument with extensions .fods, .ods and .ots

:mag: **Compare**

Compare original and converted spreadsheets to log differences.[^3]
* Cell values

:file_cabinet: **Archive**

The program can convert, package and describe spreadsheets to meet a data quality level, that will enable you to open your spreadsheets many years from now. 
* Convert any spreadsheet[^1] to both .xlsx (Strict conformance) and .ods
* Package spreadsheets and metadata in a new archive directory
* Output all conversions in subdirectories named n+1
* Rename all conversions 1.xlsx and 1.ods
* Include copies of the original spreadsheets, this include password protected or otherwise unreadable files
* Validate spreadsheet against its file format standard (Office Open XML)
* Check if spreadsheet contains any sheets and any cell values
* Remove formula linking cells to other local spreadsheets but keep cell values
* Remove external data connections but keep cell values
* Remove RealTimeData (RTD) functions but keep cell values
* Remove printer settings (feature not working)
* Alert if spreadsheet has embedded or external objects
* Make first sheet the active sheet
* Calculate file checksums
* Zip the archive directory

## Dependencies
Prerequisite software for the program to work with these functions.

:warning: **Convert**
* [LibreOffice](https://www.libreoffice.org/)
  - If you want to convert OpenDocument spreadsheets and/or use the archiving method
  - You need to install program in its default directory
  - The program is free

:warning: **Compare**
* [Beyond Compare 4](https://www.scootersoftware.com/)
  - If you want to use the compare function
  - You need to install program in its default directory
  - You need to purchase license, trial period is 30 days
  
:warning: **Archive**
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)
  - If you want to use the archiving method, which converts .xlsx conformance from Transitional to Strict
  - You need to install program in its default directory
  - You need to purchase

## How to use
Download the executable version [here](https://github.com/Asbjoedt/CLISC/releases). There's no need to install. In your terminal change directory to the folder where CLISC.exe is. Then, to execute the program input:

```
.\CLISC.exe [your_arguments]
```

Create your arguments from the following list:

**Functions to use** (required, pick one of the four)
```
--function count
--function countconvert
--function countconvertcompare
--function countconvertcomparearchive
```
**Input directory** (required)
```
--inputdir "[path to input directory]"
```
**Output directory** (required)
```
--outputdir "[path to output directory]"
```
**Include subdirectories from input directory** (optional, by default false)
```
--recurse true
```
**Example of full usage**
```
.\CLISC.exe --function countconvertcomparearchive --inputdir "c:\my_path" --outputdir "c:\my_path" --recurse true
```
**or shorter**
```
.\CLISC.exe -f countconvertcomparearchive -i "c:\my_path" -o "c:\my_path" -r true
```

If you want to test the application, a sample dataset is provided [here](https://github.com/Asbjoedt/CLISC/tree/master/Sample%20data).

## Packages and software

The following packages and software are used under license in CLISC. [Read more](https://github.com/Asbjoedt/CLISC/wiki/Dependencies).

* [Beyond Compare 4](https://www.scootersoftware.com/index.php), Copyright (c) 2022 Scooter Software, Inc.
* [CommandLineParser](https://github.com/commandlineparser/commandline), MIT License, Copyright (c) 2005 - 2015 Giacomo Stelluti Scala & Contributor
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation

[^1]: See definition of [accepted spreadsheet file formats](https://github.com/Asbjoedt/CLISC/wiki/Spreadsheet-File-Formats).
[^2]: Conversion to file extension .xlsx with Strict conformance is currently not supported.
[^3]: The program can currently not compare cell formatting, embedded objects, charts and other advanced spreadsheet features.
