using System;
using System.IO;
//using CLISC.Classes;

namespace CLISC
{

    class Program
    {

        // User guidance
        public void Main(string[] args)
        {
            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count, Convert & Compare");
            Console.WriteLine();
            Console.WriteLine("Program behavior:");
            Console.WriteLine("\tCount spreadsheets in directory by file format");
            Console.WriteLine("\tConvert XLS, XLT, XLAM, XLSB, XLTX, XLSM, XLTM to XLSX (OOXML Transitional conformance)");
            Console.WriteLine("\tOutput all conversions in same directory");
            Console.WriteLine("\tRename all conversions n+1.xlsx");
            Console.WriteLine("\tCompare the results to log workbook differences between input and output file formats");
            Console.WriteLine();
            Console.WriteLine("Available arguments:");
            Console.WriteLine("\t[value] | Filepath to directory e.g. C:\\Users\\[your_username]\\Desktop | Mandatory");
            Console.WriteLine("\trecursive | Recursively count & convert spreadsheets in any subdirectories | Optional (not working)");
            Console.WriteLine("\tprefix=[value] | Prefix filename i.e. [value]n+1.xlsx | Optional (not working)");
            Console.WriteLine("\tnolog | Output no XML log | Optional (not working)");
            Console.WriteLine();
            Console.WriteLine("Input your argument:");
        }
        
        // User input
        public argument(string directory, string recursive = "", string prefix ="", string nolog = "")
        {
            Console.ReadLine();

            //Validate user input
            if ()
            if (recursive != "recursive" || recursive != "")
            {
                Console.WriteLine("Invalid recursive argument");
            }
            if (prefix != "prefix=" + || prefix != "")
            {
                Console.WriteLine("Invalid prefix argument");
            }
            if (nolog != "nolog" || nolog != "")
            {
                Console.WriteLine("Invalid nolog argument");
            }
            // Directory will be validated in the next step
            //while (IsWellFormedUriString == false)
            //{
            //    if (public static bool IsWellFormedUriString(string? argument, UriKind 0))
            //        {
            //        IsWellFormedUriString = true;
            //        }
            //    else
            //        {
            //        Console.WriteLine("Please enter a valid filepath");
            //        argument = Console.ReadLine();
            //        }
            //}
        }

        // Count
        public Count()
        { 
            DirectoryInfo di = new DirectoryInfo(@directory);
            int numXLS = di.GetFiles("*.xls", SearchOption.AllDirectories).Length;
            int numXLT = di.GetFiles("*.xlt", SearchOption.AllDirectories).Length;
            int numXLAM = di.GetFiles("*.xlam", SearchOption.AllDirectories).Length;
            int numXLSB = di.GetFiles("*.xlsb", SearchOption.AllDirectories).Length;
            int numXLSM = di.GetFiles("*.xlsm", SearchOption.AllDirectories).Length;
            int numXLSX = di.GetFiles("*.xlsx", SearchOption.AllDirectories).Length;
            int numXLTM = di.GetFiles("*.xltm", SearchOption.AllDirectories).Length;
            int numXLTX = di.GetFiles("*.xltx", SearchOption.AllDirectories).Length;
            // Show count to user
            Console.WriteLine();
            Console.WriteLine($"{numXLS} XLS");
            Console.WriteLine($"{numXLT} XLT");
            Console.WriteLine($"{numXLAM} XLAM");
            Console.WriteLine($"{numXLSB} XLSB");
            Console.WriteLine($"{numXLSM} XLSM");
            Console.WriteLine($"{numXLSX} XLSX");
            Console.WriteLine($"{numXLTM} XLTM");
            Console.WriteLine($"{numXLTX} XLTX");
            Console.WriteLine();
                // If no spreadsheets identified ask for new argument
                if (numXLS == 0 && numXLT == 0 && numXLAM == 0 && numXLSB == 0 && numXLSM == 0 && numXLSX == 0 && numXLTM == 0 && numXLTX == 0)
                {
                    Console.WriteLine("No spreadsheets identified. Input new argument:");
                    directory = Console.ReadLine();
                }
        }
        
        // Convert
        Console.WriteLine("Continue to conversion y/n");
            // linje nedenunder bør være bool i stedet for string
            string continue_conversion = Console.ReadLine();
            if (continue_conversion == "y")
            {
                Console.WriteLine("funktion på vej");
                //Insert convert function
            }
            else
            {
                Environment.Exit(0);
            }

            // Rename
            // int filenumber = 1;
            // if (prefix has value)
            // {
            // filename = prefix + ++filenumber + ".xlsx"
            // }
            // else 
            // filename = ++filenumber + ".xlsx"
            // Compare

            // Log

            // User confirmation
            //Console.WriteLine("Conversion completed");
            //Console.WriteLine("X ""conversions contain differences");
            //Console.WriteLine("Comparison results saved to log");
    }
}