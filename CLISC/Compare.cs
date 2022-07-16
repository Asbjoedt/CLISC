using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace CLISC
{

    public partial class Spreadsheet
    {

        // Compare spreadsheets
        public void Compare(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath,Original filesize,Original checksum,New convert filepath, New filesize,New convert cheksum");
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
            string conv_filename = "1.xlsx";
            int numTOTAL_conv = 0;
            foreach (var folder in folder_enumeration)
            {
                
                Console.WriteLine(folder);
            }

            if (conv_filename == "1.xlsx")
            {
                numTOTAL_conv++;

                try
                {
                    // Foreach folder with conversion in enumeration

                    //Create "Beyond Compare" script file
                    string bcscript_filename = results_directory + "\\bcscript.txt";
                    string compare_org_filepath = "???";
                    string compare_conv_filepath = "???";

                    using (StreamWriter bcscript = File.CreateText(bcscript_filename))
                    {
                        bcscript.WriteLine(compare_org_filepath);
                        bcscript.WriteLine(compare_conv_filepath);
                    }

                    // Calculate checksums
                    int? original_checksum = null;
                    int? conv_checksum = null;

                    // Find filesizes
                    int? original_filesize = null;
                    int? conv_filesize = null;

                    // Delete BC script
                    //File.Delete(bcscript_filename);

                    // Author: Kamil Niklasinski
                    // This script is provided under GNU license -see license file for details.
                    // Make sure you add to system path folder with SPREADSHEETCOMPARE.EXE
                    // C:\Program Files(x86)\Microsoft Office\Office15\DCF\

                    //excomp.bat Book1.xlsx Book2.xlsx
                    //dir % 1 / B / S > temp.txt
                    //dir % 2 / B / S >> temp.txt
                    //SPREADSHEETCOMPARE temp.txt

                    // Log


                    // Output result in open CSV file
                    var newLine1 = string.Format($"{compare_org_filepath}, {original_filesize},{original_checksum},{compare_conv_filepath},{conv_filesize},{conv_checksum}");
                    csv.AppendLine(newLine1);




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

            else
            {
                
            }

            // Close CSV file to log results
            string convert_CSV_filepath = results_directory + "\\3_Compare_Results.csv";
            File.WriteAllText(convert_CSV_filepath, csv.ToString());

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine($"{numTOTAL_conv} converted spreadsheets were compared");
            //Console.WriteLine($"{} out of {numTOTAL_conv} conversions have workbook differences");
            Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Comparison finished");
            Console.WriteLine("---");

        }

    }

}
