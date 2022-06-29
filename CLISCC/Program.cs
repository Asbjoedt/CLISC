using System;
using System.IO;

namespace CLISC
{

    class Program
    {

        public static void Main(string[] args)
        {
            // User guidance
            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count, Convert & Compare");
            Console.WriteLine();
            Console.WriteLine("Program behavior:");
            Console.WriteLine("\tCount spreadsheets in directory by file format");
            Console.WriteLine("\tConvert XLS, XLT, XLAM, XLSB, XLTX, XLSM, XLTM to XLSX (OOXML Transitional conformance)");
            Console.WriteLine("\tOutput all conversions in a new directory with new subdirectories named n+1");
            Console.WriteLine("\tRename all conversions n+1.xlsx");
            Console.WriteLine("\tCompare the results to log workbook differences between input and output file formats");
            Console.WriteLine("\tOutput log in CSV");
            Console.WriteLine();
            Console.WriteLine("Available arguments:");
            Console.WriteLine("\t[value] | Filepath to directory e.g. C:\\Users\\[your_username]\\Desktop | Mandatory");
            Console.WriteLine("\trecursive | Recursively count & convert spreadsheets in any subdirectories | Optional");
            Console.WriteLine("\tprefix=[value] | Prefix filename i.e. [value]n+1.xlsx | Optional");
            Console.WriteLine();
            // New object reference
            Spreadsheets spreadsheet = new Spreadsheets();
            // Method references
            Spreadsheets.UserInput();
            //Spreadsheets.Count();
            //Spreadsheets.ConfirmConversion();
            //Spreadsheets.Convert();
            //Spreadsheets.Compare();
            // User confirmation
            //Console.WriteLine($"{} out of {numTOTAL} conversions completed");
            //Console.WriteLine($"{} out of {numTOTAL} conversions have differences");
            //Console.WriteLine("Results saved to log in CSV file format");
        }

    }

}