# CLISC
**Command Line Interface Spreadsheet Count, Convert, Compare & Archive**

A Windows console application made in C#. It is a prototype project for digital archiving of spreadsheets.

* For more information, see the **[wiki](https://github.com/Asbjoedt/CLISC/wiki)**
* For graphical user interface, see repository **[GUISC](https://github.com/Asbjoedt/GUISC)**
* For C# library, see repository **[Archsheerary](https://github.com/Asbjoedt/Archsheerary)**
* For simple archival workflow conversion tool, see repository **[convert-spreadsheet](https://github.com/Asbjoedt/convert-spreadsheet)**
* For simple archival workflow validation tool, see repository **[validate-spreadsheet](https://github.com/Asbjoedt/validate-spreadsheet)**
* For OpenDocument Spreadsheets and Apache POI tool, see repository **[ODS-ArchivalRequirements](https://github.com/Asbjoedt/ODS-ArchivalRequirements)**

:rainbow_flag: **General**

* Batch convert spreadsheets in a directory to .xlsx
* Include or exclude subdirectories recursively
* Output results in a new directory with logs in .csv

:heavy_plus_sign: **Count**

Count number of spreadsheets in directory by file format. 
* Accepted file extensions: .gsheet, .fods, .numbers, .ods, .ots, .xla, .xlam, .xls, .xlsb, .xlsm, .xlsx, .xlt, .xltm, .xltx
* .xlsx of Transitional and Strict conformance can be counted separately

:magic_wand: **Convert**

Convert any spreadsheet[^1][^2] to .xlsx (Transitional conformance).
* Office Open XML (Excel) with extensions .xlsb, .xlsm, .xltm, .xltx and .xlsx with Strict conformance
* Legacy Microsoft Excel with extensions .xls and .xlt
* OpenDocument with extensions .fods, .ods and .ots
* Apple Numbers with extension .numbers

:mag: **Compare**

Compare original and converted spreadsheets to log differences.[^3]
* Cell values

:file_cabinet: **Archive**

The program can convert, package and describe spreadsheets to meet a data quality level, that will enable you to open your spreadsheets many years from now. 
* Convert any spreadsheet[^1][^2] to both .xlsx (Strict conformance) and .ods
* Package spreadsheets and metadata in a new archive directory
* Output all conversions in subdirectories named n+1
* Rename all conversions 1.xlsx and 1.ods
* Include copies of the original spreadsheets, this include password protected or otherwise unreadable files
* Validate spreadsheet against its file format standard (Office Open XML and OpenDocument)
* Check if cell values exist
* Remove cell references to other spreadsheets but keep cell values
* Remove data connections but keep cell values
* Remove RealTimeData (RTD) functions but keep cell values
* Remove printer settings
* Remove absolute path to local directory
* Embed external objects (work in progress)
* Convert embedded images to .tiff
* Make first sheet active
* Alert if metadata detected
* Alert if hyperlinks detected
* Calculate file checksums
* Zip the archive directory

## Dependencies

:warning: **[Beyond Compare 4](https://www.scootersoftware.com/)**
* If you want to use the compare function
* You need to install program in its default directory, or create environment variable "BeyondCompare" with path to your installation

:warning: **[LibreOffice](https://www.libreoffice.org/)**
* If you want to convert OpenDocument spreadsheets and/or use the archiving method
* You need to install program in its default directory, or create environment variable "LibreOffice" with path to your installation

:warning: **[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)**
* If you want to convert legacy Excel and/or use the archiving method, which converts .xlsx conformance from Transitional to Strict

:warning: **[ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html)**
* If you want to validate .ods spreadsheets
* You need to install program in "C:\Program Files\ODF Validator" and name program "odfvalidator-0.10.0-jar-with-dependencies.jar", or create environment variable "ODFValidator" with path to your installation
* ODF Validator needs latest version of Java Development Kit installed

## How to use
Download the executable version [here](https://github.com/Asbjoedt/CLISC/releases). There's no need to install. In your terminal change directory to the folder where CLISC.exe is. Then, to execute the program input:

```
.\CLISC.exe [your_arguments]
```

Create your arguments from the following list:

**Functions to use** (required, pick one of the four)
```
--function Count
--function CountConvert
--function CountConvertCompare
--function CountConvertCompareArchive
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
--recurse
```
**Example of full usage**
```
.\CLISC.exe --function CountConvertCompareArchive --inputdir "c:\folder" --outputdir "c:\folder" --recurse
```
**or shorter**
```
.\CLISC.exe -f CountConvertCompareArchive -i "c:\folder" -o "c:\folder" -r
```

If you want to test the application, a sample dataset is provided [here](https://github.com/Asbjoedt/CLISC/blob/master/Docs/SampleData.zip).

## Packages and software

The following packages and software are used under license in CLISC. [Read more](https://github.com/Asbjoedt/CLISC/wiki/Dependencies).

* [Beyond Compare 4](https://www.scootersoftware.com/index.php), copyright (c) 2022 Scooter Software, Inc.
* [CommandLineParser](https://github.com/commandlineparser/commandline), MIT License, copyright (c) 2005 - 2015 Giacomo Stelluti Scala & Contributor
* [Magick.Net](https://github.com/dlemstra/Magick.NET), Apache-2.0 license, copyright (c) dlemstra
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), copyright (c) Microsoft Corporation
* [ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html), Apache License, [copyright info](https://github.com/tdf/odftoolkit/blob/master/NOTICE)
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, copyright (c) Microsoft Corporation

[^1]: See definition of [accepted spreadsheet file formats](https://github.com/Asbjoedt/CLISC/wiki/Spreadsheet-File-Formats).
[^2]: The program currently has a conversion filesize limit of 150MB to prevent excessive performance bottlenecks. Larger filesize spreadsheets should be converted manually.
[^3]: The program can currently not compare cell formatting, embedded objects, charts and other advanced spreadsheet features.
