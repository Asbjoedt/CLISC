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
        public void Count(string argument1, string argument2, string argument3)
        {

            // Variables
            int numODS, numOTS, numFODS, numXLS, numXLT, numXLAM, numXLSB, numXLSM, numXLSX, numXLTM, numXLTX, numTOTAL;

            // Arrays
            string[] file_format = { "Extension", "FODS", "ODS", "OTS", "XLS", "XLT", "XLAM", "XLSB", "XLSM", "XLSX", "XLTM", "XLTX" };
            string[] file_format_description = { "Name", "OpenDocument Flat XML Spreadsheet", "OpenDocument Spreadsheet", "OpenDocument Spreadsheet Template", "Legacy Microsoft Excel Spreadsheet", "Legacy Microsoft Excel Spreadsheet Template", "Office Open XML Macro-Enabled Add-In", "Office Open XML Binary Spreadsheet", "Office Open XML Macro-Enabled Spreadsheet", "Office Open XML Spreadsheet", "Office Open XML Macro-Enabled Spreadsheet Template", "Office Open XML Spreadsheet Template" };

            //Object reference
            DirectoryInfo dir = new DirectoryInfo(argument1);

            // Spreadsheets to count
            if (argument3 == "Recursive=Yes")
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
            else if (argument3 == "Recursive=No")
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
            else
            {
                Console.WriteLine("Invalid argument in position args[3]");
            }

            // Inform user if no spreadsheets identified
            if (numFODS == 0 && numODS == 0 && numOTS == 0 && numXLS == 0 && numXLT == 0 && numXLAM == 0 && numXLSB == 0 && numXLSM == 0 && numXLSX == 0 && numXLTM == 0 && numXLTX == 0)
            {
                Console.WriteLine("No spreadsheets identified.");
                Console.WriteLine();
            }
            else
            {
                // Show count to user
                Console.WriteLine($"# {file_format[0]} - {file_format_description[0]}");
                Console.WriteLine($"{numFODS} {file_format[1]} - {file_format_description[1]}");
                Console.WriteLine($"{numODS} {file_format[2]} - {file_format_description[2]}");
                Console.WriteLine($"{numOTS} {file_format[3]} - {file_format_description[3]}");
                Console.WriteLine($"{numXLS} {file_format[4]} - {file_format_description[4]}");
                Console.WriteLine($"{numXLT} {file_format[5]} - {file_format_description[5]}");
                Console.WriteLine($"{numXLAM} {file_format[6]} - {file_format_description[6]}");
                Console.WriteLine($"{numXLSB} {file_format[7]} - {file_format_description[7]}");
                Console.WriteLine($"{numXLSM} {file_format[8]} - {file_format_description[8]}");
                Console.WriteLine($"{numXLSX} {file_format[9]} - {file_format_description[9]}");
                Console.WriteLine($"{numXLTM} {file_format[10]} - {file_format_description[10]}");
                Console.WriteLine($"{numXLTX} {file_format[11]} - {file_format_description[11]}");
                Console.WriteLine($"{numTOTAL} spreadsheets in total");

                // Output results in CSV
                var csv = new StringBuilder();
                var newLine0 = string.Format($"#,{file_format[0]},{file_format_description[0]}");
                csv.AppendLine(newLine0);
                var newLine1 = string.Format($"{numFODS},{file_format[1]},{file_format_description[1]}");
                csv.AppendLine(newLine1);
                var newLine2 = string.Format($"{numODS},{file_format[2]},{file_format_description[2]}");
                csv.AppendLine(newLine2);
                var newLine3 = string.Format($"{numOTS},{file_format[3]},{file_format_description[3]}");
                csv.AppendLine(newLine3);
                var newLine4 = string.Format($"{numXLS},{file_format[4]},{file_format_description[4]}");
                csv.AppendLine(newLine4);
                var newLine5 = string.Format($"{numXLT},{file_format[5]},{file_format_description[5]}");
                csv.AppendLine(newLine5);
                var newLine6 = string.Format($"{numXLAM},{file_format[6]},{file_format_description[6]}");
                csv.AppendLine(newLine6);
                var newLine7 = string.Format($"{numXLSB},{file_format[7]},{file_format_description[7]}");
                csv.AppendLine(newLine7);
                var newLine8 = string.Format($"{numXLSM},{file_format[8]},{file_format_description[8]}");
                csv.AppendLine(newLine8);
                var newLine9 = string.Format($"{numXLSX},{file_format[9]},{file_format_description[9]}");
                csv.AppendLine(newLine9);
                var newLine10 = string.Format($"{numXLTM},{file_format[10]},{file_format_description[10]}");
                csv.AppendLine(newLine10);
                var newLine11 = string.Format($"{numXLTX},{file_format[11]},{file_format_description[11]}");
                csv.AppendLine(newLine11);
                File.WriteAllText(argument2, csv.ToString());
                Console.WriteLine("Results saved to log in CSV file format");
                // Inform user of end of Count
            }
            Console.WriteLine("Count finished");
            Console.WriteLine();

        }

        // Convert spreadsheets
        public void Convert(string argument1, string argument2, string argument3)
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
        public void Compare(string argument1, string argument2, string argument3)
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
