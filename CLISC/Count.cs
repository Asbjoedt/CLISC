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
        // Public integers
        public int numTOTAL, numXLSX_Strict;

        // Count spreadsheets
        public string Count(string argument1, string argument2, string argument3)
        {
            Console.WriteLine("COUNT");
            Console.WriteLine("---");

            // Integers
            int numFODS, numODS, numOTS, numXLA, numXLAM, numXLS, numXLSB, numXLSM, numXLSX, numXLSX_Transitional, numXLT, numXLTM, numXLTX;

            //Object reference
            DirectoryInfo count = new DirectoryInfo(argument1);
            FileFormats info = new FileFormats();

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
                numXLSX_Transitional = numXLSX - (numXLSX_Strict + numCONFORM_fail);
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
                numXLSX_Transitional = numXLSX - (numXLSX_Strict + numCONFORM_fail);
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
                Console.WriteLine($"{numFODS} {FileFormats.Extension[0]} - {FileFormats.Description[0]}");
                Console.WriteLine($"{numODS} {FileFormats.Extension[1]} - {FileFormats.Description[1]}");
                Console.WriteLine($"{numOTS} {FileFormats.Extension[2]} - {FileFormats.Description[2]}");
                Console.WriteLine($"{numXLA} {FileFormats.Extension[3]} - {FileFormats.Description[3]}");
                Console.WriteLine($"{numXLAM} {FileFormats.Extension[4]} - {FileFormats.Description[4]}");
                Console.WriteLine($"{numXLS} {FileFormats.Extension[5]} - {FileFormats.Description[5]}");
                Console.WriteLine($"{numXLSB} {FileFormats.Extension[6]} - {FileFormats.Description[6]}");
                Console.WriteLine($"{numXLSM} {FileFormats.Extension[7]} - {FileFormats.Description[7]}");
                Console.WriteLine($"{numXLSX} {FileFormats.Extension[8]} - {FileFormats.Description[8]}");
                Console.WriteLine($"--> {numXLSX_Transitional} of {numXLSX} {FileFormats.Extension[8]} have Office Open XML Transitional conformance");
                Console.WriteLine($"--> {numXLSX_Strict} of {numXLSX} {FileFormats.Extension[8]} have Office Open XML Strict conformance");
                if (numCONFORM_fail > 0) 
                {
                        Console.WriteLine($"--> {numCONFORM_fail} of {numXLSX} {FileFormats.Extension[8]} could not be identified because they are password protected or corrupt");
                }
                Console.WriteLine($"{numXLT} {FileFormats.Extension[9]} - {FileFormats.Description[9]}");
                Console.WriteLine($"{numXLTM} {FileFormats.Extension[10]} - {FileFormats.Description[10]}");
                Console.WriteLine($"{numXLTX} {FileFormats.Extension[11]} - {FileFormats.Description[11]}");

                // Create new directory to output results in CSV
                string Results_Directory = Create_Directory_Results(argument1, argument2);

                // Output results in CSV
                var csv = new StringBuilder();
                var newLine0 = string.Format("#;Extension;Name");
                csv.AppendLine(newLine0);
                var newLine1 = string.Format($"{numFODS};{FileFormats.Extension[0]};{FileFormats.Description[0]}");
                csv.AppendLine(newLine1);
                var newLine2 = string.Format($"{numODS};{FileFormats.Extension[1]};{FileFormats.Description[1]}");
                csv.AppendLine(newLine2);
                var newLine3 = string.Format($"{numOTS};{FileFormats.Extension[2]};{FileFormats.Description[2]}");
                csv.AppendLine(newLine3);
                var newLine4 = string.Format($"{numXLA};{FileFormats.Extension[3]};{FileFormats.Description[3]}");
                csv.AppendLine(newLine4);
                var newLine5 = string.Format($"{numXLAM};{FileFormats.Extension[4]};{FileFormats.Description[4]}");
                csv.AppendLine(newLine5);
                var newLine6 = string.Format($"{numXLS};{FileFormats.Extension[5]};{FileFormats.Description[5]}");
                csv.AppendLine(newLine6);
                var newLine7 = string.Format($"{numXLSB};{FileFormats.Extension[6]};{FileFormats.Description[6]}");
                csv.AppendLine(newLine7);
                var newLine8 = string.Format($"{numXLSM};{FileFormats.Extension[7]};{FileFormats.Description[7]}");
                csv.AppendLine(newLine8);
                var newLine9 = string.Format($"{numXLSX};{FileFormats.Extension[8]};{FileFormats.Description[8]}");
                csv.AppendLine(newLine9);
                var newLine13 = string.Format($"{numXLSX_Transitional};.xlsx Transitional;Office Open XML Spreadsheet Transitional conformance");
                csv.AppendLine(newLine13);
                var newLine14 = string.Format($"{numXLSX_Strict};.xlsx Strict;Office Open XML Spreadsheet Strict conformance");
                csv.AppendLine(newLine14);
                var newLine15 = string.Format($"{numCONFORM_fail};.xlsx unknown;Conformance could not be identified");
                csv.AppendLine(newLine15);
                var newLine10 = string.Format($"{numXLT};{FileFormats.Extension[9]};{FileFormats.Description[9]}");
                csv.AppendLine(newLine10);
                var newLine11 = string.Format($"{numXLTM};{FileFormats.Extension[10]};{FileFormats.Description[10]}");
                csv.AppendLine(newLine11);
                var newLine12 = string.Format($"{numXLTX};{FileFormats.Extension[11]};{FileFormats.Description[11]}");
                csv.AppendLine(newLine12);
                string CSV_filepath = Results_Directory + "\\1_Count_Results.csv";
                File.WriteAllText(CSV_filepath, csv.ToString());

                // Inform user of results
                Console.WriteLine("---");
                Console.WriteLine($"{numTOTAL} spreadsheet files in total");
                Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
                Console.WriteLine("Count finished");
                Console.WriteLine("---");

                return Results_Directory;
            }
        }
    }
}
