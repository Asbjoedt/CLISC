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
        public void Compare(string argument1, string argument2, string argument3, string argument4)
        {

            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filesize (KB);Conversion identified;Conversion filepath;Conversion filesize (KB)");
            csv.AppendLine(newLine0);

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
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Comparing to: {compare_conv_filepath}");

                    // Compare workbook differences
                    if (File.Exists(compare_conv_filepath))
                    {
                        compare_success = true;
                        numTOTAL_conv++;

                        // Compare workbooks using external app Beyond Compare 4
                        Compare_Workbook(results_directory, folder, compare_org_filepath, compare_conv_filepath);

                        // Calculate filesize of converted spreadsheet
                        conv_filesize_kb = Calculate_Filesize(compare_conv_filepath);

                    }

                    // Calculate filesize of original spreadsheet
                    org_filesize_kb = Calculate_Filesize(compare_org_filepath);

                    // Output result in open CSV file
                    var newLine1 = string.Format($"{org_filepath};{org_filesize_kb};{compare_success};{compare_conv_filepath};{conv_filesize_kb};");
                    csv.AppendLine(newLine1);

                }

                // Delete BC script
                if (File.Exists(bcscript_filepath))
                {
                    File.Delete(bcscript_filepath);
                }

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
