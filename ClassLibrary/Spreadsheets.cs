using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace CLISC
{

    public class Spreadsheets
    {
        // Declare public variables
        public string directory;
        public string prefix = ""; // = null; ???
        //public string valid_prefix = prefix.Length=>8;

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
                Console.WriteLine($"{recursive}: Subdirectories will be included");
            }
            else if (recursiveString == "false")
            {
                Console.WriteLine($"{recursive}: Subdirectories will not be included");
            }
            else
            {
                Console.WriteLine("Input not valid");
                // Restart method or create another kind of loop?
            }
            Console.WriteLine();
            // Use prefix
            Console.WriteLine("Input prefix for renaming conversions. If no prefix hit 'Enter' without writing anything");
            prefix = Console.ReadLine();
            //return (directory, recursive, prefix);
        }

        // Count spreadsheets
        public void Count()
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
            int numTOTAL = numXLS + numXLT + numXLAM + numXLSB + numXLSM + numXLSX + numXLTM + numXLTX;
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
            Console.WriteLine($"{numTOTAL} spreadsheets in total");
            Console.WriteLine();
            // If no spreadsheets identified, ask for new argument
            if (numXLS == 0 && numXLT == 0 && numXLAM == 0 && numXLSB == 0 && numXLSM == 0 && numXLSX == 0 && numXLTM == 0 && numXLTX == 0)
            {
                Console.WriteLine("No spreadsheets identified. Input new argument:");
                directory = Console.ReadLine();
            }
        }
        public void ConfirmConversion()
        {
            Console.WriteLine("Continue to conversion? y/n");
            string continue_conversion = Console.ReadLine();
            if (continue_conversion == "y")
            {
                Console.WriteLine("Funktion på vej");
            }
            else
            {
                Environment.Exit(0);
            }
        }

        // Convert spreadsheets
        public void Convert()
        {

            // Create new folder, copy and rename spreadsheet
            // createfolder_copyfile_renamefile()
            // Rename
            // int filenumber = 1;
            // if (prefix has value)
            // {
            // filename = prefix + ++filenumber + ".xlsx"
            // }
            // else 
            // filename = ++filenumber + ".xlsx"

            // Convert spreadsheet
        }

        // Compare spreadsheets
        public void Compare()
        {
            Console.WriteLine("funktion på vej");
            //Delete copy
            // Log
        }

    }

}
