using System;
using System.IO;
using System.Collections.Generic;
using CommandLine;

namespace CLISC
{
    public class Program_Real
    {
        public static void Execute(string function, string inputdir, string outputdir, bool recurse)
        {
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
                    app.Final_Results();
                    break;

                case "count&convert":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    con.Convert_Spreadsheets(inputdir, recurse, resultsDirectory);
                    app.Final_Results();
                    break;

                case "count&convert&compare":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    List<fileIndex> fileList = con.Convert_Spreadsheets(inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    app.Final_Results();
                    break;

                case "count&convert&compare&archive":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    fileList = con.Convert_Spreadsheets_Archive(inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    arc.Archive_Spreadsheets(resultsDirectory, fileList);
                    app.Final_Results();
                    break;

                default:
                    Console.WriteLine("Invalid function argument. Function argument must be one these: count, count&convert, count&convert&compare, count&convert&compare&archive");
                    break;
            }
        }

        public void Final_Results()
        {
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheets");
            Console.WriteLine($"CONVERT: {Conversion.numCOMPLETE} out of {Count.numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"CONVERT: {Conversion.numXLSX_noconversion} spreadsheets were already .xlsx");
            Console.WriteLine($"CONVERT: {Conversion.numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"COMPARE: {Compare.numTOTAL_compare} out of {Conversion.numTOTAL_conv} converted spreadsheets were compared");
            Console.WriteLine($"ARCHIVE: {Archive.valid_files} out of {Conversion.numTOTAL_conv} converted spreadsheets have valid file formats");
            Console.WriteLine($"ARCHIVE: {Archive.extrels_files} out of {Conversion.numTOTAL_conv} converted spreadsheets had external relationships - External relationships were removed");
            Console.WriteLine($"ARCHIVE: {Archive.embedobj_files} out of {Conversion.numTOTAL_conv} converted spreadsheets have embedded objects - Embedded objects were NOT removed");
            Console.WriteLine("CLISC ended");
            Console.WriteLine("---");
        }
    }
}