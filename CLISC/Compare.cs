using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{

    public partial class Spreadsheet
    {

        // Compare spreadsheets
        public void Compare(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Comparison errors
            bool success;
            int numTOTAL_diff = 0;

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filesize;Original checksum;New convert filepath;New filesize;New convert cheksum");
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

            // Identify existence of converted spreadsheet
            string compare_org_filepath = "orgFile_" + "";
            string compare_conv_filepath = docCollection + "\\1.xlsx";
            int numTOTAL_conv = 0;

            foreach (var folder in folder_enumeration)
            {
                File.Exists(compare_conv_filepath);

                // Calculate checksums
                var org_checksum = CalculateMD5(compare_org_filepath);
                var conv_checksum = CalculateMD5(compare_conv_filepath);

                // Find filesizes
                long filesize;

                FileInfo fi = new FileInfo(compare_org_filepath);
                filesize = fi.Length;
                long original_filesize = filesize;

                new FileInfo(compare_conv_filepath);
                filesize = fi.Length;
                long conv_filesize = filesize;

                // Compare workbook differences
                if (File.Exists(compare_conv_filepath))
                {
                    numTOTAL_conv++;

                    try
                    {
                        //Create "Beyond Compare" script file
                        string bcscript_filename = results_directory + "\\bcscript.txt";
                        using (StreamWriter bcscript = File.CreateText(bcscript_filename))
                        {
                            bcscript.WriteLine(compare_org_filepath);
                            bcscript.WriteLine(compare_conv_filepath);
                        }

                        // Use BC

                        // .\Path to\BComp.com

                        // Delete BC script
                        File.Delete(bcscript_filename);

                        // If there is workbook differences
                        //if (fail)
                        //{
                        //    numTOTAL_diff++;
                        //
                        //    // Inform user
                        //    Console.WriteLine(compare_conv_filepath);
                        //    Console.WriteLine($"--> Comparison {success} - Workbook differences identified");
                        //}

                        // No workbook differences
                        //else
                        //{
                        //    // Inform user
                        //    Console.WriteLine(compare_conv_filepath);
                        //    Console.WriteLine($"--> Comparison {success}");
                        //}

                    }

                    // Error message if BC is not detected
                    catch (FileNotFoundException)
                    {
                        Console.WriteLine("Error: The program Microsoft Spreadsheet Compare is necessary for compare function to run.");
                    }

                    finally
                    {

                    }

                }

                // Output result in open CSV file
                var newLine1 = string.Format($"{compare_org_filepath};{original_filesize};{org_checksum};{compare_conv_filepath};{conv_filesize};{conv_checksum}");
                csv.AppendLine(newLine1);

            }

            // Close CSV file to log results
            string convert_CSV_filepath = results_directory + "\\3_Compare_Results.csv";
            File.WriteAllText(convert_CSV_filepath, csv.ToString());

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine($"{numTOTAL_conv} conversions out of {numTOTAL} spreadsheets were compared");
            //Console.WriteLine($"{numTOTAL_diff} out of {numTOTAL_conv} conversions have workbook differences");
            Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Comparison finished");
            Console.WriteLine("---");

        }

    }

}
