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
    public partial class Spreadsheet
    {
        // Comparison data types
        public static int numTOTAL_conv = 0;
        public static int numTOTAL_compare = numTOTAL_conv;
        public static int numTOTAL_diff = 0;
        string compare_message = "";

        // Compare spreadsheets
        public void Compare(string Results_Directory, List<fileIndex> File_List)
        {
            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filesize (KB);Conversion filepath;Conversion filesize (KB);Filesize diff;Workbook diff");
            csv.AppendLine(newLine0);

            if (File.Exists(@"C:\\Program Files\\Beyond Compare 4\\BCompare.exe"))
            {
                try
                {
                    foreach (fileIndex entry in File_List)
                    {
                        // Get information from list
                        string org_filepath = entry.Org_Filepath;
                        string conv_filepath = entry.Conv_Filepath;
                        string folder = entry.File_Folder;

                        // Compare workbook differences
                        if (File.Exists(conv_filepath))
                        {
                            numTOTAL_conv++;

                            // Inform user of comparison
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Comparing to: {conv_filepath}");

                            // Compare workbooks using external app Beyond Compare 4
                            Compare_Workbook(Results_Directory, folder, org_filepath, conv_filepath);

                            // Calculate filesize of converted spreadsheet
                            int conv_filesize_kb = Calculate_Filesize(conv_filepath);

                            // Calculate filesize of original spreadsheet
                            int org_filesize_kb = Calculate_Filesize(org_filepath);

                            // File size diff
                            bool filesize_diff;
                            if (conv_filesize_kb == org_filesize_kb)
                            {
                                filesize_diff = true;
                            }
                            else
                            {
                                filesize_diff = false;
                            }

                            // Output result in open CSV file
                            var newLine1 = string.Format($"{org_filepath};{org_filesize_kb};{conv_filepath};{conv_filesize_kb};{filesize_diff};{compare_message}");
                            csv.AppendLine(newLine1);
                        }
                    }
                }
                // Error message if BC is not detected
                catch (Win32Exception)
                {
                    compare_message = "Beyond Compare 4 is not installed in filepath: C:\\Program Files\\Beyond Compare 4";
                }
            }
            else
            {
                Console.WriteLine("Beyond Compare 4 is not installed in filepath: C:\\Program Files\\Beyond Compare 4");
                Console.WriteLine("Comparison ended");
            }

            // Close CSV file to log results
            string CSV_filepath = Results_Directory + "\\3_Compare_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Inform user of results
            Compare_Results();
        }

        public void Compare_Results()
        {
            string CSV_filepath = Results_Directory + "\\3_Compare_Results.csv";

            Console.WriteLine("---");
            Console.WriteLine("COMPARE RESULTS");
            Console.WriteLine("---");
            Console.WriteLine($"{numTOTAL_compare} out of {numTOTAL_conv} converted spreadsheets were compared");
            //Console.WriteLine($"{numTOTAL_diff} out of {numTOTAL_conv} conversions have workbook differences");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
            Console.WriteLine("Comparison ended");
            Console.WriteLine("---");
        }
    }
}
