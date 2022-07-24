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
        public int numFODS, numODS, numOTS, numXLA, numXLS, numXLT, numXLAM, numXLSB, numXLSM, numXLSX, numXLSX_Strict, numXLSX_Transitional, numXLTM, numXLTX, numTOTAL;

        // Arrays
        public string[] File_Format = { ".fods", ".ods", ".ots", ".xla", ".xlam", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltm", ".xltx" };

        public string[] File_Format_Description = { "OpenDocument Flat XML Spreadsheet", "OpenDocument Spreadsheet", "OpenDocument Spreadsheet Template", "Legacy Microsoft Excel Spreadsheet Add-In", "Legacy Microsoft Excel Spreadsheet", "Legacy Microsoft Excel Spreadsheet Template", "Office Open XML Macro-Enabled Add-In", "Office Open XML Binary Spreadsheet", "Office Open XML Macro-Enabled Spreadsheet", "Office Open XML Spreadsheet (Transitional and Strict conformance)", "Office Open XML Macro-Enabled Spreadsheet Template", "Office Open XML Spreadsheet Template" };

        // Count spreadsheets
        public string Count(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("COUNT");
            Console.WriteLine("---");

            //Object reference
            DirectoryInfo count = new DirectoryInfo(argument1);

            // Count spreadsheets recursively
            if (argument3 == "Recurse=Yes")
            {
                numFODS = count.GetFiles("*.fods", SearchOption.AllDirectories).Length;
                numODS = count.GetFiles("*.ods", SearchOption.AllDirectories).Length;
                numOTS = count.GetFiles("*.ots", SearchOption.AllDirectories).Length;
                numXLA = count.GetFiles("*.xla", SearchOption.AllDirectories).Length;
                numXLS = count.GetFiles("*.xls", SearchOption.AllDirectories).Length;
                numXLT = count.GetFiles("*.xlt", SearchOption.AllDirectories).Length;
                numXLAM = count.GetFiles("*.xlam", SearchOption.AllDirectories).Length;
                numXLSB = count.GetFiles("*.xlsb", SearchOption.AllDirectories).Length;
                numXLSM = count.GetFiles("*.xlsm", SearchOption.AllDirectories).Length;
                numXLSX = count.GetFiles("*.xlsx", SearchOption.AllDirectories).Length;
                numXLSX_Strict = Count_XLSX_Strict(argument1, argument3);
                numXLSX_Transitional = numXLSX - (numXLSX_Strict + conformance_count_fail);
                numXLTM = count.GetFiles("*.xltm", SearchOption.AllDirectories).Length;
                numXLTX = count.GetFiles("*.xltx", SearchOption.AllDirectories).Length;

                numTOTAL = numFODS + numODS + numOTS + numXLA + numXLS + numXLT + numXLAM + numXLSB + numXLSM + numXLSX + numXLTM + numXLTX;
            }

            // Count spreadsheets NOT recursively
            else
            {
                numFODS = count.GetFiles("*.fods", SearchOption.TopDirectoryOnly).Length;
                numODS = count.GetFiles("*.ods", SearchOption.TopDirectoryOnly).Length;
                numOTS = count.GetFiles("*.ots", SearchOption.TopDirectoryOnly).Length;
                numXLA = count.GetFiles("*.xla", SearchOption.TopDirectoryOnly).Length;
                numXLS = count.GetFiles("*.xls", SearchOption.TopDirectoryOnly).Length;
                numXLT = count.GetFiles("*.xlt", SearchOption.TopDirectoryOnly).Length;
                numXLAM = count.GetFiles("*.xlam", SearchOption.TopDirectoryOnly).Length;
                numXLSB = count.GetFiles("*.xlsb", SearchOption.TopDirectoryOnly).Length;
                numXLSM = count.GetFiles("*.xlsm", SearchOption.TopDirectoryOnly).Length;
                numXLSX = count.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly).Length;
                numXLSX_Strict = Count_XLSX_Strict(argument1, argument3);
                numXLSX_Transitional = numXLSX - (numXLSX_Strict + conformance_count_fail);
                numXLTM = count.GetFiles("*.xltm", SearchOption.TopDirectoryOnly).Length;
                numXLTX = count.GetFiles("*.xltx", SearchOption.TopDirectoryOnly).Length;

                numTOTAL = numFODS + numODS + numOTS + numXLA + numXLS + numXLT + numXLAM + numXLSB + numXLSM + numXLSX + numXLTM + numXLTX;
            }

            // Inform user if no spreadsheets identified
            if (numTOTAL == 0)
            {
                Console.WriteLine("No spreadsheets identified");
                Console.WriteLine("Count finished");
                Console.WriteLine("---");

                throw new Exception();
            }
            else
            {
                // Show count to user
                Console.WriteLine("# Extension - Name");
                Console.WriteLine($"{numFODS} {File_Format[0]} - {File_Format_Description[0]}");
                Console.WriteLine($"{numODS} {File_Format[1]} - {File_Format_Description[1]}");
                Console.WriteLine($"{numOTS} {File_Format[2]} - {File_Format_Description[2]}");
                Console.WriteLine($"{numXLA} {File_Format[3]} - {File_Format_Description[3]}");
                Console.WriteLine($"{numXLAM} {File_Format[4]} - {File_Format_Description[4]}");
                Console.WriteLine($"{numXLS} {File_Format[5]} - {File_Format_Description[5]}");
                Console.WriteLine($"{numXLSB} {File_Format[6]} - {File_Format_Description[6]}");
                Console.WriteLine($"{numXLSM} {File_Format[7]} - {File_Format_Description[7]}");
                Console.WriteLine($"{numXLSX} {File_Format[8]} - {File_Format_Description[8]}");
                Console.WriteLine($"--> {numXLSX_Transitional} of {numXLSX} {File_Format[8]} have Office Open XML Transitional conformance");
                Console.WriteLine($"--> {numXLSX_Strict} of {numXLSX} {File_Format[8]} have Office Open XML Strict conformance");
                if (conformance_count_fail > 0) 
                {
                        Console.WriteLine($"--> {conformance_count_fail} of {numXLSX} {File_Format[8]} could not be counted because they are password protected or corrupt");
                }
                Console.WriteLine($"{numXLT} {File_Format[9]} - {File_Format_Description[9]}");
                Console.WriteLine($"{numXLTM} {File_Format[10]} - {File_Format_Description[10]}");
                Console.WriteLine($"{numXLTX} {File_Format[11]} - {File_Format_Description[11]}");

                // Create new directory to output results in CSV
                Results_Directory = Create_Directory_Results(argument1, argument2);

                // Output results in CSV
                var csv = new StringBuilder();
                var newLine0 = string.Format("#;Extension;Name");
                csv.AppendLine(newLine0);
                var newLine1 = string.Format($"{numFODS};{File_Format[0]};{File_Format_Description[0]}");
                csv.AppendLine(newLine1);
                var newLine2 = string.Format($"{numODS};{File_Format[1]};{File_Format_Description[1]}");
                csv.AppendLine(newLine2);
                var newLine3 = string.Format($"{numOTS};{File_Format[2]};{File_Format_Description[2]}");
                csv.AppendLine(newLine3);
                var newLine4 = string.Format($"{numXLA};{File_Format[3]};{File_Format_Description[3]}");
                csv.AppendLine(newLine4);
                var newLine5 = string.Format($"{numXLAM};{File_Format[4]};{File_Format_Description[4]}");
                csv.AppendLine(newLine5);
                var newLine6 = string.Format($"{numXLS};{File_Format[5]};{File_Format_Description[5]}");
                csv.AppendLine(newLine6);
                var newLine7 = string.Format($"{numXLSB};{File_Format[6]};{File_Format_Description[7]}");
                csv.AppendLine(newLine7);
                var newLine8 = string.Format($"{numXLSM};{File_Format[7]};{File_Format_Description[7]}");
                csv.AppendLine(newLine8);
                var newLine9 = string.Format($"{numXLSX};{File_Format[8]};{File_Format_Description[8]}");
                csv.AppendLine(newLine9);
                var newLine10 = string.Format($"{numXLT};{File_Format[9]};{File_Format_Description[9]}");
                csv.AppendLine(newLine10);
                var newLine11 = string.Format($"{numXLTM};{File_Format[10]};{File_Format_Description[10]}");
                csv.AppendLine(newLine11);
                var newLine12 = string.Format($"{numXLTX};{File_Format[11]};{File_Format_Description[11]}");
                csv.AppendLine(newLine12);
                string CSV_filepath = Results_Directory + "\\1_Count_Results.csv";
                File.WriteAllText(CSV_filepath, csv.ToString());

                // Inform user of results
                Console.WriteLine("---");
                Console.WriteLine($"{numTOTAL} spreadsheet files in total");
                Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
                Console.WriteLine("Count finished");
                Console.WriteLine("---");

                return results_directory;

            }

        }

    }

}
