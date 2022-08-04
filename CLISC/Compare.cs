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
        public void Compare_Spreadsheets(string function, string Results_Directory, List<fileIndex> File_List)
        {
            Console.WriteLine("---");
            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Data types
            string error_message = "Beyond Compare 4 is not installed in filepath: C:\\Program Files\\Beyond Compare 4";
            int org_filesize_kb;
            int xlsx_filesize_kb;
            int ods_filesize_kb;
            bool xlsx_filesize_diff;
            bool ods_filesize_diff;
            var xlsx_compare_message = "";
            var ods_compare_message = "";

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;Original Filesize (KB);XLSX Convert Filepath;XLSX Filesize (KB);XLSX Filesize Diff;XLSX Compare Message; ODS Convert Filepath; ODS Filesize (KB); ODS Filesize Diff; ODS Compare Message");
            csv.AppendLine(newLine0);

            if (File.Exists(@"C:\\Program Files\\Beyond Compare 4\\BCompare.exe"))
            {

                foreach (fileIndex entry in File_List)
                {
                    // Get information from list
                    string org_filepath = entry.Org_Filepath;
                    string xlsx_filepath = entry.XLSX_Conv_Filepath;
                    string ods_filepath = entry.ODS_Conv_Filepath;
                    string folder = entry.File_Folder;

                    // Compare workbook differences
                    if (File.Exists(xlsx_filepath))
                    {
                        numTOTAL_compare++;

                        if (function == "count&convert&compare&archive")
                        {
                            // Compare workbooks using external app Beyond Compare 4
                            xlsx_compare_message = Compare_Workbook(Results_Directory, folder, org_filepath, xlsx_filepath);
                            //ods_compare_message = Compare_Workbook(Results_Directory, folder, org_filepath, ods_filepath);

                            // Calculate filesizes
                            org_filesize_kb = Calculate_Filesize(org_filepath);
                            xlsx_filesize_kb = Calculate_Filesize(xlsx_filepath);

                            ods_filesize_kb = Calculate_Filesize(ods_filepath);

                            // Determine file size diff
                            if (xlsx_filesize_kb == org_filesize_kb)
                            {
                                xlsx_filesize_diff = true;
                            }
                            else
                            {
                                xlsx_filesize_diff = false;
                            }

                            if (ods_filesize_kb == org_filesize_kb)
                            {
                                ods_filesize_diff = true;
                            }
                            else
                            {
                                ods_filesize_diff = false;
                            }

                            // Inform user of comparison
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Comparing to: {xlsx_filepath}");
                            Console.WriteLine(xlsx_compare_message);

                            // Output result in open CSV file
                            var newLine1 = string.Format($"{org_filepath};{org_filesize_kb};{xlsx_filepath};{xlsx_filesize_kb};{xlsx_filesize_diff};{xlsx_compare_message};{ods_filepath};{ods_filesize_kb};{ods_filesize_diff};{ods_compare_message}");
                            csv.AppendLine(newLine1);
                        }

                        // No archiving
                        else
                        {
                            // Compare workbooks using external app Beyond Compare 4
                            xlsx_compare_message = Compare_Workbook(Results_Directory, folder, org_filepath, xlsx_filepath);

                            // Calculate filesizes
                            org_filesize_kb = Calculate_Filesize(org_filepath);
                            xlsx_filesize_kb = Calculate_Filesize(xlsx_filepath);

                            // Determine file size diff
                            if (xlsx_filesize_kb == org_filesize_kb)
                            {
                                xlsx_filesize_diff = true;
                            }
                            else
                            {
                                xlsx_filesize_diff = false;
                            }

                            // Inform user of comparison
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Comparing to: {xlsx_filepath}");
                            Console.WriteLine(xlsx_compare_message);

                            // Output result in open CSV file - BUG becasue string.Format cannot handle list output from BC4
                            //var newLine2 = string.Format($"{org_filepath};{org_filesize_kb};{xlsx_filepath};{xlsx_filesize_kb};{xlsx_filesize_diff};{xlsx_compare_message};;;;");
                            //csv.AppendLine(newLine2);
                        }
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
