using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    public partial class Compare
    {
        // Comparison data types
        public static int numTOTAL_compare = 0;
        public static int numTOTAL_diff = 0;

        // Compare spreadsheets
        public void Compare_Spreadsheets(string Results_Directory, List<fileIndex> File_List)
        {
            Console.WriteLine("---");
            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Data types
            string error_message = "Beyond Compare 4 is not installed in filepath: C:\\Program Files\\Beyond Compare 4";

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;XLSX Filepath;Comparison Success");
            csv.AppendLine(newLine0);

            if (File.Exists(@"C:\\Program Files\\Beyond Compare 4\\BCompare.exe"))
            {

                foreach (fileIndex entry in File_List)
                {
                    bool compare_success = false;
                    // Get information from list
                    string org_filepath = entry.Org_Filepath;
                    string xlsx_filepath = entry.XLSX_Conv_Filepath;
                    string ods_filepath = entry.ODS_Conv_Filepath;
                    string folder = entry.File_Folder;

                    // Compare workbook differences
                    if (File.Exists(xlsx_filepath))
                    {
                        // Compare workbooks using external app Beyond Compare 4
                        int return_code = Compare_Workbook(Results_Directory, folder, org_filepath, xlsx_filepath);

                        // Inform user of comparison
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Comparing to: {xlsx_filepath}");
                        if (return_code == 0 || return_code == 1 || return_code == 2)
                        {
                            numTOTAL_compare++;
                            compare_success = true;
                            Console.WriteLine("--> Spreadsheets identical: " + compare_success);
                        }
                        if (return_code == 12 || return_code == 13 || return_code == 14)
                        {
                            numTOTAL_compare++;
                            numTOTAL_diff++;
                            compare_success = false;
                            Console.WriteLine("--> Spreadsheets identical: " + compare_success);
                        }
                        if (return_code == 11)
                        {
                            compare_success = false;
                            Console.WriteLine("--> Original file is a .fods, .ods, .ots or .xlsb spreadsheet and cannot be compared");
                        }

                        // Output result in open CSV file
                        var newLine1 = string.Format($"{org_filepath};{xlsx_filepath};{compare_success}");
                        csv.AppendLine(newLine1);
                    }
                }
            }
            else
            {
                Console.WriteLine(error_message);
                Console.WriteLine("Comparison ended");
            }

            // Close CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\3_Compare_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());
        }
    }
}
