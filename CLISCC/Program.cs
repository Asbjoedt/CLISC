using System;
//using CLISSC.Classes;

namespace CLISSC
{
    class Program
    {
        static void Main(string[] args)
        {
            // Declare variables
            string argument1, prefix;
            bool recursive, nolog;
            // User guidance
            Console.WriteLine("CLISCC - Command Line Interface Spreadsheet Convert and Compare");
            Console.WriteLine();
            Console.WriteLine("Program behavior:");
            Console.WriteLine("\tConvert XLS, XLT, XLAM, XLSB, XLTX, XLSM, XLTM to XLSX (OOXML Transitional conformance)");
            Console.WriteLine("\tOutput all conversions in same folder");
            Console.WriteLine("\tRename all conversions n+1.xlsx");
            Console.WriteLine("\tCompare the results to log workbook differences between input and output file formats");
            Console.WriteLine();
            Console.WriteLine("Available arguments:");
            Console.WriteLine("\t[value] | Filepath to folder e.g. C:\\Users\\[your_username]\\Desktop | Mandatory");
            Console.WriteLine("\tprefix=[value] | Prefix filename i.e. [value]n+1.xlsx | Optional (not working)");
            Console.WriteLine("\trecursive | Recursively convert spreadsheets in any subfolders | Optional (not working)");
            Console.WriteLine("\tnolog | Output no XML log | Optional (not working)");
            Console.WriteLine();
            Console.WriteLine("Input your argument:");
            // User input
            argument1 = Console.ReadLine();
            //Validate user input
            while (public static bool IsWellFormedUriString (string? argument1, UriKind 0))
            {
                //Perform functions
                break
                Console.WriteLine("Please enter a valid filepath");
                argument1 = Console.ReadLine();
            }
            
            // Functions
            // - Convert

            // - Compare

            // - Log

            // User confirmation
            //Console.WriteLine("Conversion completed");
            //Console.WriteLine("X ""conversions contain differences");
            //Console.WriteLine("Comparison results saved to log");
        }
    }
}