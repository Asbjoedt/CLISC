using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC.Classes
{
    internal class Test
    {
        //Declare public variables
        public string directory;

        // User guidance
        static void UserGuidance()
        {   
            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count, Convert & Compare");
            Console.WriteLine();
            Console.WriteLine("Program behavior:");
            Console.WriteLine("\tCount Excel spreadsheets in directory by file format");
            Console.WriteLine("\tConvert XLS, XLT, XLAM, XLSB, XLTX, XLSM, XLTM to XLSX (OOXML Transitional conformance)");
            Console.WriteLine("\tOutput all conversions in a new directory with new subdirectories named n+1");
            Console.WriteLine("\tRename all conversions n+1.xlsx");
            Console.WriteLine("\tCompare the results to log workbook and checksum differences between input and output file formats");
            Console.WriteLine("\tOutput log in CSV");
            Console.WriteLine();
            Console.WriteLine("Available arguments:");
            Console.WriteLine("\tCount 'Filepath to directory' -Recursive");
            Console.WriteLine("\tCount&Convert 'Filepath to directory' -Recursive");
            Console.WriteLine("\tCount&Convert&Compare 'Filepath to directory' -Recursive");
            Console.WriteLine();
            Console.WriteLine("Input your argument:");
        }

        // User input
        public void UserInput()
        {
            // Input directory
            Console.WriteLine("Input directory path:");
            directory = Console.ReadLine();
            Console.WriteLine();
            // Include subdirectories
            Console.WriteLine("Include subdirectories? Input 'true' or 'false'");
            string recursiveString = Console.ReadLine();
            bool recursive = recursiveString == "true";
            if (recursiveString == "true")
            {
                Console.WriteLine("Subdirectories will be included");
            }
            else if (recursiveString == "false")
            {
                Console.WriteLine("Subdirectories will be excluded");
            }
            else
            {
                Console.WriteLine("Input not valid");
                // Restart method or create another kind of loop?
            }
            //return (directory);
        }

        // User confirmation prompt
        static void Confirm()
        {
            Console.WriteLine("Continue to next process y/n");
            string continue_conversion = Console.ReadLine();
            if (continue_conversion == "y")
            {
                Console.WriteLine();
                Console.WriteLine("Funktion på vej");
            }
            else
            {
                Environment.Exit(0);
            }
        }
    }
}
