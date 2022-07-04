using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        
        // Public variables
        public int numFODS, numODS, numOTS, numXLS, numXLT, numXLAM, numXLSB, numXLSM, numXLSX, numXLTM, numXLTX, numTOTAL;

        // Count spreadsheets
        public void Count(string argument1, string argument2, string argument3)
        {

            // Arrays
            string[] file_format = { "Extension", ".fods", ".ods", ".ots", ".xls", ".xlt", ".xlam", ".xlsb", ".xlsm", ".xlsx", ".xltm", ".xltx" };

            string[] file_format_description = { "Name", "OpenDocument Flat XML Spreadsheet", "OpenDocument Spreadsheet", "OpenDocument Spreadsheet Template", "Legacy Microsoft Excel Spreadsheet", "Legacy Microsoft Excel Spreadsheet Template", "Office Open XML Macro-Enabled Add-In", "Office Open XML Binary Spreadsheet", "Office Open XML Macro-Enabled Spreadsheet", "Office Open XML Spreadsheet", "Office Open XML Macro-Enabled Spreadsheet Template", "Office Open XML Spreadsheet Template" };

            //Object reference
            DirectoryInfo InputDir = new DirectoryInfo(argument1);

            // Spreadsheets to count
            if (argument3 == "Recursive=Yes")
            {
                numFODS = InputDir.GetFiles("*.fods", SearchOption.AllDirectories).Length;
                numODS = InputDir.GetFiles("*.ods", SearchOption.AllDirectories).Length;
                numOTS = InputDir.GetFiles("*.ots", SearchOption.AllDirectories).Length;
                numXLS = InputDir.GetFiles("*.xls", SearchOption.AllDirectories).Length;
                numXLT = InputDir.GetFiles("*.xlt", SearchOption.AllDirectories).Length;
                numXLAM = InputDir.GetFiles("*.xlam", SearchOption.AllDirectories).Length;
                numXLSB = InputDir.GetFiles("*.xlsb", SearchOption.AllDirectories).Length;
                numXLSM = InputDir.GetFiles("*.xlsm", SearchOption.AllDirectories).Length;
                numXLSX = InputDir.GetFiles("*.xlsx", SearchOption.AllDirectories).Length;
                numXLTM = InputDir.GetFiles("*.xltm", SearchOption.AllDirectories).Length;
                numXLTX = InputDir.GetFiles("*.xltx", SearchOption.AllDirectories).Length;
                numTOTAL = numFODS + numODS + numOTS + numXLS + numXLT + numXLAM + numXLSB + numXLSM + numXLSX + numXLTM + numXLTX;
            }
            else if (argument3 == "Recursive=No")
            {
                numFODS = InputDir.GetFiles("*.fods", SearchOption.TopDirectoryOnly).Length;
                numODS = InputDir.GetFiles("*.ods", SearchOption.TopDirectoryOnly).Length;
                numOTS = InputDir.GetFiles("*.ots", SearchOption.TopDirectoryOnly).Length;
                numXLS = InputDir.GetFiles("*.xls", SearchOption.TopDirectoryOnly).Length;
                numXLT = InputDir.GetFiles("*.xlt", SearchOption.TopDirectoryOnly).Length;
                numXLAM = InputDir.GetFiles("*.xlam", SearchOption.TopDirectoryOnly).Length;
                numXLSB = InputDir.GetFiles("*.xlsb", SearchOption.TopDirectoryOnly).Length;
                numXLSM = InputDir.GetFiles("*.xlsm", SearchOption.TopDirectoryOnly).Length;
                numXLSX = InputDir.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly).Length;
                numXLTM = InputDir.GetFiles("*.xltm", SearchOption.TopDirectoryOnly).Length;
                numXLTX = InputDir.GetFiles("*.xltx", SearchOption.TopDirectoryOnly).Length;
                numTOTAL = numFODS + numODS + numOTS + numXLS + numXLT + numXLAM + numXLSB + numXLSM + numXLSX + numXLTM + numXLTX;
            }
            else
            {
                Console.WriteLine("Invalid argument in position args[3]");
            }

            // Inform user if no spreadsheets identified
            if (numTOTAL == 0)
            {
                Console.WriteLine("No spreadsheets identified");
                Environment.Exit(0);
            }
            else
            {
                // Show count to user
                Console.WriteLine("Count");
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

                // Create new directory to output results in CSV
                int results_directory_number = 1;
                string results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
                while (Directory.Exists(@results_directory))
                {
                    results_directory_number++;
                    results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
                }
                DirectoryInfo OutputDir = Directory.CreateDirectory(@results_directory);

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
                string count_CSV_filepath = results_directory + "\\1_Count_Results.csv";
                File.WriteAllText(count_CSV_filepath, csv.ToString());
                Console.WriteLine($"Results saved to CSV log in filepath: {count_CSV_filepath}");
            }

            // Inform user of end of Count method
            Console.WriteLine("Count finished");
            Console.WriteLine("---");


        }

    }

}
