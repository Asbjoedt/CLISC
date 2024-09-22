using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;

namespace CLISC
{
    public class Program_Real
    {
        public static void Execute(string function, string inputdir, string outputdir, bool recurse, bool fullcompliance)
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
            Results res = new Results();

            // Method references
            switch (function)
            {
                case "Count":
                    cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    res.Count_Results();
                    break;

                case "CountConvert":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    con.Convert_Spreadsheets(function, inputdir, recurse, resultsDirectory);
                    res.Convert_Results();
                    break;

                case "CountConvertCompare":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    List<fileIndex> fileList = con.Convert_Spreadsheets(function, inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    res.Compare_Results();
                    break;

                case "CountConvertCompareArchive":
                    resultsDirectory = cou.Count_Spreadsheets(inputdir, outputdir, recurse);
                    fileList = con.Convert_Spreadsheets(function, inputdir, recurse, resultsDirectory);
                    com.Compare_Spreadsheets(function, resultsDirectory, fileList);
                    arc.Archive_Spreadsheets(resultsDirectory, fileList, fullcompliance);
                    res.Archive_Results(fullcompliance);
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
            Console.WriteLine("CLISC complete");
            Console.WriteLine("---"); 
        }
    }
}