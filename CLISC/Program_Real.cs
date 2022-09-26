using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using CommandLine;

namespace CLISC
{
    public class Program_Real
    {
        public static void Execute(string function, string inputdir, string outputdir, bool recurse)
        {
            if (!Directory.Exists(inputdir))
            {
                throw new DirectoryNotFoundException("Input directory does not exist");
            }

            // Begin process timer
            Stopwatch timer = new Stopwatch();
            timer.Start();

            // Path to new directory in output directory
            string resultsDirectory = "";

            // Create object instance
            Count cou = new Count();
            Conversion con = new Conversion();
            Compare com = new Compare();
            Archive arc = new Archive();
            Program_Real app = new Program_Real();

            // Method references
            switch (function)
            {
                case "Count":
                    cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    app.Count_Results();
                    break;

                case "CountConvert":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    con.Convert_Spreadsheets(function, inputdir, recurse, resultsDirectory);
                    app.Convert_Results();
                    break;

                case "CountConvertCompare":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    List<fileIndex> fileList = con.Convert_Spreadsheets(function, inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    app.Compare_Results();
                    break;

                case "CountConvertCompareArchive":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    fileList = con.Convert_Spreadsheets(function, inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    arc.Archive_Spreadsheets(resultsDirectory, fileList);
                    app.Archive_Results();
                    break;

                default:
                    Console.WriteLine("Invalid function argument. Function argument must be one these: Count, CountConvert, CountConvertCompare, CountConvertCompareArchive");
                    break;
            }

            // Stop process timer
            timer.Stop();
            TimeSpan time = timer.Elapsed;
            string elapsedTime = String.Format($"{time:dd\\:hh\\:mm\\:ss} (days:hrs:min:sec)");
            Console.WriteLine("Total process time: " + elapsedTime);
            Console.WriteLine("CLISC ended");
            Console.WriteLine("---"); 
        }

        // Methods for results reporting
        void Count_Results()
        {
            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"Results saved to CSV log in filepath: {Spreadsheet.CSV_filepath}");
        }

        void Convert_Results()
        {
            // Calculate fails
            int fail_conversion = Count.numTOTAL - Conversion.numCOMPLETE;

            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"CONVERT: {fail_conversion} of {Count.numTOTAL} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {Spreadsheet.CSV_filepath}");
        }
        void Compare_Results()
        {
            // Calculate fails
            int fail_conversion = Count.numTOTAL - Conversion.numCOMPLETE;
            int fail_comparison = Conversion.numCOMPLETE - Compare.numTOTAL_compare;

            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"CONVERT: {fail_conversion} of {Count.numTOTAL} spreadsheets failed conversion");
            Console.WriteLine($"COMPARE: {fail_comparison} of {Conversion.numCOMPLETE} converted spreadsheets failed comparison");
            Console.WriteLine($"COMPARE: {Compare.numTOTAL_diff} of {Compare.numTOTAL_compare} compared spreadsheets have cell value differences");

        }
        void Archive_Results()
        {
            // Calculate fails
            int fail_conversion = Count.numTOTAL - Conversion.numCOMPLETE;
            int fail_comparison = Conversion.numCOMPLETE - Compare.numTOTAL_compare;

            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheets");
            Console.WriteLine($"CONVERT: {fail_conversion} of {Count.numTOTAL} spreadsheets failed conversion");
            Console.WriteLine($"COMPARE: {fail_comparison} of {Conversion.numCOMPLETE} converted spreadsheets failed comparison");
            Console.WriteLine($"COMPARE: {Compare.numTOTAL_diff} of {Compare.numTOTAL_compare} compared spreadsheets have cell value differences");
            Console.WriteLine($"ARCHIVE: {Archive.invalid_files} of {Conversion.numCOMPLETE} converted spreadsheets have invalid file formats");
            Console.WriteLine($"ARCHIVE: {Archive.cellvalue_files} of {Conversion.numCOMPLETE} converted spreadsheets had no cell values - Handle manually!");
            Console.WriteLine($"ARCHIVE: {Archive.connections_files} of {Conversion.numCOMPLETE} converted spreadsheets had data connections - Data connections were removed");
            Console.WriteLine($"ARCHIVE: {Archive.cellreferences_files} of {Conversion.numCOMPLETE} converted spreadsheets had external cell references - External cell references were removed");
            Console.WriteLine($"ARCHIVE: {Archive.rtdfunctions_files} of {Conversion.numCOMPLETE} converted spreadsheets had RTD functions - RTD functions were removed");
            Console.WriteLine($"ARCHIVE: {Archive.extobj_files} of {Conversion.numCOMPLETE} converted spreadsheets had external object references - External object references were removed");
            Console.WriteLine($"ARCHIVE: {Archive.embedobj_files} of {Conversion.numCOMPLETE} converted spreadsheets had embedded objects  - Embedded objects were removed");
            Console.WriteLine($"ARCHIVE: {Archive.printersettings_files} of {Conversion.numCOMPLETE} converted spreadsheets had printer settings - Printer settings were removed");
            Console.WriteLine($"ARCHIVE: {Archive.activesheet_files} of {Conversion.numCOMPLETE} converted spreadsheets did not have first sheet as active sheet - Active sheet was changed");
        }
    }
}