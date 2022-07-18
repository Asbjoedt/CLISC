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
        bool compare_success = false;
        int numTOTAL_diff = 0;
        string conv_checksum = "";
        string org_checksum = "";
        int? org_filesize_kb = null;
        int? conv_filesize_kb = null;
        string[] compare_error_message = { "", "Beyond Compare 4 is not installed in filepath: C:\\Program Files\\Beyond Compare 4" };

        // Compare spreadsheets
        public void Compare(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filesize (KB);Original checksum;Conversion identified;Conversion filepath;Conversion filesize (KB);Conversion checksum");
            csv.AppendLine(newLine0);

            // Identify CLISC subdirectory
            int results_directory_number = 1;
            string results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
            while (Directory.Exists(@results_directory))
            {
                results_directory_number++;
                results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
            }
            results_directory_number = results_directory_number - 1;
            results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;

            // Identify docCollection
            string docCollection = results_directory + "\\docCollection";

            // Create enumeration of docCollection
            var folder_enumeration = new FileSystemEnumerable<string>(
                docCollection,
                (ref FileSystemEntry entry) => entry.ToFullPath(),
                new EnumerationOptions() { RecurseSubdirectories = false })
            {
                ShouldIncludePredicate = (ref FileSystemEntry entry) => entry.IsDirectory
            };

            // Loop through docCollection enumeration
            string compare_org_filepath = "";
            string compare_conv_filepath = "";
            int numTOTAL_conv = 0;

            foreach (var folder in folder_enumeration)
            {

                // Identify original file in folder
                var org_file = from file in
                Directory.EnumerateFiles(folder)
                            where file.Contains("orgFile_") // Should this be similar code to line 81?
                            select file;
                foreach (var file in org_file)
                {
                    compare_org_filepath = file.ToString();
                }

                // Identify converted spreadsheet in folder
                var conv_file = from file in
                Directory.EnumerateFiles(folder)
                            where Path.GetFileName(file).Equals("1.xlsx")
                            select file;
                foreach (var file in conv_file)
                {
                    compare_conv_filepath = file.ToString();

                    // Inform user of comparison
                    Console.WriteLine(compare_org_filepath);
                    Console.WriteLine($"--> Comparing to: {compare_conv_filepath}");

                    // Compare workbook differences
                    if (File.Exists(compare_conv_filepath))
                    {
                        compare_success = true;
                        numTOTAL_conv++;

                        // Compare workbooks using external app Beyond Compare 4
                        Compare_Workbook(results_directory, folder, compare_org_filepath, compare_conv_filepath);

                        // Calculate MD5 of converted spreadsheet
                        conv_checksum = Calculate_MD5(compare_conv_filepath);

                        // Calculate filesize of converted spreadsheet
                        conv_filesize_kb = Calculate_Filesize(compare_conv_filepath);

                    }

                    // Calculate checksum of original spreadsheet
                    org_checksum = Calculate_MD5(compare_org_filepath);

                    // Calculate filesize of original spreadsheet
                    org_filesize_kb = Calculate_Filesize(compare_org_filepath);

                    // Output result in open CSV file
                    var newLine1 = string.Format($"{compare_org_filepath};{org_filesize_kb};{org_checksum};{compare_success};{compare_conv_filepath};{conv_filesize_kb};{conv_checksum}");
                    csv.AppendLine(newLine1);

                }

            }

            // Close CSV file to log results
            string convert_CSV_filepath = results_directory + "\\3_Compare_Results.csv";
            File.WriteAllText(convert_CSV_filepath, csv.ToString());

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine($"{numTOTAL_conv} out of {numTOTAL} spreadsheets were compared");
            //Console.WriteLine($"{numTOTAL_diff} out of {numTOTAL_conv} conversions have workbook differences");
            Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Comparison finished");
            Console.WriteLine("---");

        }

    }

}
