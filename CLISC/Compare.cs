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
        public int numTOTAL_diff = 0;
        string compare_message = "Beyond Compare 4 is not installed in filepath: C:\\Program Files\\Beyond Compare 4";

        // Compare spreadsheets
        public void Compare(string argument0, string argument1, string results_directory, List<string> docCollection_enumeration)
        {

            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            int numTOTAL_conv = 0;

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filesize (KB);Conversion filepath;Conversion filesize (KB); Comparison message");
            csv.AppendLine(newLine0);

            if (argument0 == "Count&Convert&Compare&Archive")
            {
                // Loop through docCollection enumeration
                foreach (var folder in docCollection_enumeration)
                {
                    // Identify converted spreadsheet in folder
                    var conv_file = from file in
                    Directory.EnumerateFiles(folder)
                                    where Path.GetFileName(file).Equals("1.xlsx")
                                    select file;
                    try
                    {
                        foreach (var file in conv_file)
                        {
                            conv_filepath = file.ToString();

                            // Compare workbook differences
                            if (File.Exists(conv_filepath))
                            {
                                numTOTAL_conv++;

                                // Inform user of comparison
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Comparing to: {conv_filepath}");

                                // Compare workbooks using external app Beyond Compare 4
                                Compare_Workbook(argument0, results_directory, folder, org_filepath, conv_filepath);

                                // Calculate filesize of converted spreadsheet
                                int conv_filesize_kb = Calculate_Filesize(conv_filepath);

                                // Calculate filesize of original spreadsheet
                                int org_filesize_kb = Calculate_Filesize(org_filepath);

                                // Output result in open CSV file
                                var newLine1 = string.Format($"{org_filepath};{org_filesize_kb};{conv_filepath};{conv_filesize_kb};{compare_message}");
                                csv.AppendLine(newLine1);
                            }
                        }
                    }

                    // Error message if BC is not detected
                    catch (Win32Exception)
                    {
                        Console.WriteLine($"--> {compare_message}");
                    }
                }
            }

            // ordinary comparison, no archiving
            else
            {
                // Identify docCollection path
                string docCollection = results_directory + "\\docCollection";

                // Loop through docCollection enumeration
                foreach (var file in docCollection_enumeration)
                {
                    numTOTAL_conv++;

                    // Inform user of comparison
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Comparing to: {conv_filepath}");

                    // Create instance for finding file information
                    FileInfo file_info = new FileInfo(file);
                    conv_filepath = file_info.FullName;

                    // Compare workbook differences
                    try
                    {
                        // Compare workbooks using external app Beyond Compare 4
                        Compare_Workbook(argument0, results_directory, docCollection, org_filepath, conv_filepath);

                        // Calculate filesize of converted spreadsheet
                        int conv_filesize_kb = Calculate_Filesize(conv_filepath);

                        // Calculate filesize of original spreadsheet
                        int org_filesize_kb = Calculate_Filesize(org_filepath);

                        // Output result in open CSV file
                        var newLine1 = string.Format($"{org_filepath};{org_filesize_kb};{conv_filepath};{conv_filesize_kb};{compare_message}");
                        csv.AppendLine(newLine1);
                    }

                    // Error message if BC is not detected
                    catch (Win32Exception)
                    {
                        Console.WriteLine($"--> {compare_message}");
                    }
                }
            }

            // Delete BC script
            if (File.Exists(bcscript_filepath))
            {
                File.Delete(bcscript_filepath);
            }

            // Close CSV file to log results
            string CSV_filepath = results_directory + "\\3_Compare_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine($"{numTOTAL_conv} out of {numTOTAL} spreadsheets were compared");
            //Console.WriteLine($"{numTOTAL_diff} out of {numTOTAL_conv} conversions have workbook differences");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
            Console.WriteLine("Comparison finished");
            Console.WriteLine("---");

        }

    }

}
