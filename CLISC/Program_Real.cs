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
            string Results_Directory = "";

            // Create object instance
            Spreadsheet process = new Spreadsheet();
            Program_Real report = new Program_Real();
            Program_Args app_arg = new Program_Args();

            // Method references
            switch (function)
            {
                case "count":
                    process.Count(inputdir, outputdir, recurse);
                    report.Final_Results();
                    break;

                case "count&convert":
                    Results_Directory = process.Count(inputdir, outputdir, recurse);
                    process.Convert(function, inputdir, recurse, Results_Directory);
                    report.Final_Results();
                    break;

                case "count&convert&compare":
                    Results_Directory = process.Count(inputdir, outputdir, recurse);
                    List<fileIndex> File_List = process.Convert(function, inputdir, recurse, Results_Directory);
                    process.Compare(Results_Directory, File_List);
                    report.Final_Results();
                    break;

                case "count&convert&compare&archive":
                    Results_Directory = process.Count(inputdir, outputdir, recurse);
                    File_List = process.Convert(function, inputdir, recurse, Results_Directory);
                    process.Compare(Results_Directory, File_List);
                    process.Archive(Results_Directory, File_List);
                    report.Final_Results();
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
            Console.WriteLine($"{Spreadsheet.numTOTAL} spreadsheets counted");
            Console.WriteLine($"{Spreadsheet.numCOMPLETE} out of {Spreadsheet.numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"{Spreadsheet.numXLSX_noconversion} spreadsheets were already .xlsx");
            Console.WriteLine($"{Spreadsheet.numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"{Spreadsheet.numTOTAL_compare} out of {Spreadsheet.numTOTAL_conv} converted spreadsheets were compared");
            Console.WriteLine($"{Spreadsheet.valid_files} out of {Spreadsheet.numTOTAL_conv} converted spreadsheets have valid file formats");
            Console.WriteLine($"{Spreadsheet.extrels_files} out of {Spreadsheet.numTOTAL_conv} converted spreadsheets had external relationships - External relationships were removed");
            Console.WriteLine($"{Spreadsheet.embedobj_files} out of {Spreadsheet.numTOTAL_conv} converted spreadsheets have embedded objects - Embedded objects were NOT removed");
            Console.WriteLine("CLISC ended");
            Console.WriteLine("---");
        }
    }
}