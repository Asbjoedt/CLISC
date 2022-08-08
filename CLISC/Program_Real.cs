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
                case "count":
                    cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    app.Count_Results();
                    break;

                case "count&convert":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    con.Convert_Spreadsheets_Archive(function, inputdir, recurse, resultsDirectory);
                    app.Convert_Results();
                    break;

                case "count&convert&compare":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    List<fileIndex> fileList = con.Convert_Spreadsheets_Archive(function, inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    app.Compare_Results();
                    break;

                case "count&convert&compare&archive":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    fileList = con.Convert_Spreadsheets_Archive(function, inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    arc.Archive_Spreadsheets(resultsDirectory, fileList);
                    app.Archive_Results();
                    break;

                default:
                    Console.WriteLine("Invalid function argument. Function argument must be one these: count, count&convert, count&convert&compare, count&convert&compare&archive");
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
            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"CONVERT: {Conversion.numCOMPLETE} out of {Count.numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"CONVERT: {Conversion.numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {Spreadsheet.CSV_filepath}");
        }
        void Compare_Results()
        {
            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"CONVERT: {Conversion.numCOMPLETE} out of {Count.numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"CONVERT: {Conversion.numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"COMPARE: {Compare.numTOTAL_compare} out of {Conversion.numCOMPLETE} converted spreadsheets were compared");
            Console.WriteLine($"COMPARE: 0 converted spreadsheets failed comparison");

        }
        void Archive_Results()
        {
            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheets");
            Console.WriteLine($"CONVERT: {Conversion.numCOMPLETE} out of {Count.numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"CONVERT: {Conversion.numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"COMPARE: {Compare.numTOTAL_compare} out of {Conversion.numCOMPLETE} converted spreadsheets completed comparison");
            Console.WriteLine($"COMPARE: 0 converted spreadsheets failed comparison");
            Console.WriteLine($"ARCHIVE: {Archive.valid_files} out of {Conversion.numCOMPLETE} converted spreadsheets have valid file formats");
            Console.WriteLine($"ARCHIVE: {Archive.invalid_files} converted spreadsheets have invalid file formats");
            Console.WriteLine($"ARCHIVE: {Archive.extrels_files} out of {Conversion.numCOMPLETE} converted spreadsheets had external relationships - External relationships were removed");
            Console.WriteLine($"ARCHIVE: {Archive.embedobj_files} out of {Conversion.numCOMPLETE} converted spreadsheets have embedded objects - Embedded objects were NOT removed");
        }
    }
}