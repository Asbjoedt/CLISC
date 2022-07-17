﻿using System;
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

        // Compare spreadsheets
        public void Compare(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Comparison data types
            bool success = false;
            int numTOTAL_diff = 0;
            var conv_checksum = "";
            var org_checksum = "";
            int? org_filesize_kb = null;
            int? conv_filesize_kb = null;

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
                            where file.Contains("orgFile_")
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
                foreach (var file2 in conv_file)
                {
                    compare_conv_filepath = file2.ToString();

                    // Inform user of comparison
                    Console.WriteLine(compare_conv_filepath);
                    Console.WriteLine($"--> Comparing to: {compare_org_filepath}");

                    // Compare workbook differences
                    if (File.Exists(compare_conv_filepath))
                    {
                        success = true;
                        numTOTAL_conv++;

                        try
                        {
                            //Create "Beyond Compare" script file
                            string bcscript_filepath = results_directory + "\\bcscript.txt";
                            string bcscript_results_filepath = folder + "\\comparisonResults.html";
                            using (StreamWriter bcscript = File.CreateText(bcscript_filepath))
                            {
                                bcscript.WriteLine($"data-report layout:interleaved options:display-mismatches title:CLISC_Comparison_Results output-to:\"{bcscript_results_filepath}\" output-options:wrap-word,html-color \"{compare_org_filepath}\" \"{compare_conv_filepath}\"");
                            }

                            // Execute BC in console
                            //Process.Start($"C:\\Program Files\\Beyond Compare 4\\BCompare.exe\" \"@{bcscript_filepath}\"");

                            // Delete BC script
                            //File.Delete(bcscript_filepath);

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
                            //    Console.WriteLine($"--> Comparison {success}");
                            //}

                        }

                        // Error message if BC is not detected
                        catch (System.ComponentModel.Win32Exception)
                        {
                            Console.WriteLine("--> Beyond Compare 4 filepath not found: C:\\Program Files\\Beyond Compare 4\\BCompare.exe");
                            Console.WriteLine("--> The program Beyond Compare 4 must be installed to compare workbook differences");
                        }

                        // Perform other comparisons
                        finally
                        {
                            // Calculate MD5 of converted spreadsheet
                            conv_checksum = CalculateMD5(compare_conv_filepath);

                            // Calculate filesize of converted spreadsheet
                            conv_filesize_kb = CalculateFilesize(compare_conv_filepath);
                        }

                    }

                    // Calculate checksum of original spreadsheet
                    org_checksum = CalculateMD5(compare_org_filepath);

                    // Calculate filesize of original spreadsheet
                    org_filesize_kb = CalculateFilesize(compare_org_filepath);

                    // Output result in open CSV file
                    var newLine1 = string.Format($"{compare_org_filepath};{org_filesize_kb};{org_checksum};{success};{compare_conv_filepath};{conv_filesize_kb};{conv_checksum}");
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
