using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace CLISC
{

    public class Spreadsheet
    {

        // Count spreadsheets
        public void Count()
        {

            // Declare variables
            int numODS, numOTS, numFODS, numXLS, numXLT, numXLAM, numXLSB, numXLSM, numXLSX, numXLTM, numXLTX, numTOTAL;

            //Object reference
            DirectoryInfo dir = new DirectoryInfo(args[1]);

            // Spreadsheets to count
            if (args[2] == "-Recursive")
            {
                numFODS = dir.GetFiles("*.fods", SearchOption.AllDirectories).Length;
                numODS = dir.GetFiles("*.ods", SearchOption.AllDirectories).Length;
                numOTS = dir.GetFiles("*.ots", SearchOption.AllDirectories).Length;
                numXLS = dir.GetFiles("*.xls", SearchOption.AllDirectories).Length;
                numXLT = dir.GetFiles("*.xlt", SearchOption.AllDirectories).Length;
                numXLAM = dir.GetFiles("*.xlam", SearchOption.AllDirectories).Length;
                numXLSB = dir.GetFiles("*.xlsb", SearchOption.AllDirectories).Length;
                numXLSM = dir.GetFiles("*.xlsm", SearchOption.AllDirectories).Length;
                numXLSX = dir.GetFiles("*.xlsx", SearchOption.AllDirectories).Length;
                numXLTM = dir.GetFiles("*.xltm", SearchOption.AllDirectories).Length;
                numXLTX = dir.GetFiles("*.xltx", SearchOption.AllDirectories).Length;
                numTOTAL = numFODS + numODS + numOTS + numXLS + numXLT + numXLAM + numXLSB + numXLSM + numXLSX + numXLTM + numXLTX;
            }
            else
            {
                numFODS = dir.GetFiles("*.fods", SearchOption.TopDirectoryOnly).Length;
                numODS = dir.GetFiles("*.ods", SearchOption.TopDirectoryOnly).Length;
                numOTS = dir.GetFiles("*.ots", SearchOption.TopDirectoryOnly).Length;
                numXLS = dir.GetFiles("*.xls", SearchOption.TopDirectoryOnly).Length;
                numXLT = dir.GetFiles("*.xlt", SearchOption.TopDirectoryOnly).Length;
                numXLAM = dir.GetFiles("*.xlam", SearchOption.TopDirectoryOnly).Length;
                numXLSB = dir.GetFiles("*.xlsb", SearchOption.TopDirectoryOnly).Length;
                numXLSM = dir.GetFiles("*.xlsm", SearchOption.TopDirectoryOnly).Length;
                numXLSX = dir.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly).Length;
                numXLTM = dir.GetFiles("*.xltm", SearchOption.TopDirectoryOnly).Length;
                numXLTX = dir.GetFiles("*.xltx", SearchOption.TopDirectoryOnly).Length;
                numTOTAL = numFODS + numODS + numOTS + numXLS + numXLT + numXLAM + numXLSB + numXLSM + numXLSX + numXLTM + numXLTX;
            }

            // Show count to user
            Console.WriteLine();
            Console.WriteLine("OpenDocument formats");
            Console.WriteLine($"{numFODS} FODS (Flat XML OpenDocument Spreadsheets)");
            Console.WriteLine($"{numODS} ODS (OpenDocument Spreadsheets)");
            Console.WriteLine($"{numOTS} OTS");
            Console.WriteLine();
            Console.WriteLine("Legacy Microsoft Excel formats");
            Console.WriteLine($"{numXLS} XLS");
            Console.WriteLine($"{numXLT} XLT");
            Console.WriteLine();
            Console.WriteLine("OfficeOpen XML formats (Microsoft Excel)");
            Console.WriteLine($"{numXLAM} XLAM");
            Console.WriteLine($"{numXLSB} XLSB");
            Console.WriteLine($"{numXLSM} XLSM");
            Console.WriteLine($"{numXLSX} XLSX");
            Console.WriteLine($"{numXLTM} XLTM");
            Console.WriteLine($"{numXLTX} XLTX");
            Console.WriteLine();
            Console.WriteLine($"**{numTOTAL}** spreadsheets in total");
            //Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Count finished");
            Console.WriteLine();

            // Inform user if no spreadsheets identified
            if (numXLS == 0 && numXLT == 0 && numXLAM == 0 && numXLSB == 0 && numXLSM == 0 && numXLSX == 0 && numXLTM == 0 && numXLTX == 0)
            {
                Console.WriteLine("Count finished. No spreadsheets identified.");
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
            Console.WriteLine();
            //Console.WriteLine($"{} out of {numTOTAL} conversions completed");
            Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Conversion finished");
            Console.WriteLine();
        }

        // Compare spreadsheets
        public void Compare()
        {
            Console.WriteLine("funktion på vej");
            //Delete copy
            // Log
            Console.WriteLine();
            //Console.WriteLine($"{} out of {numTOTAL} conversions have differences");
            Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Comparison finished");
            Console.WriteLine();
        }

    }

}
