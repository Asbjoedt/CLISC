using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {

        // Archive the spreadsheets according to advanced archival requirements
        public void Archive(string argument1, string argument2, string results_directory)
        {
            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filesize (KB);Original checksum;Conversion identified;Conversion filepath;Conversion filesize (KB);Conversion checksum");
            csv.AppendLine(newLine0);

            // Rename and move converted spreadsheets


            // Copy original spreadsheets


            // Validate file format standards
            switch (file_info.Extension)
            {

                // Validate OpenDocument file formats
                case ".fods":
                case ".ods":
                case ".ots":

                    break;

                // Validate Office Open XML file formats
                case ".xlam":
                case ".xlsm":
                case ".xlsx":
                case ".xltx":
                    Validate_OOXML(argument1, argument2);
                    break;
            }

            // Calculate checksums
            string copy_checksum = Calculate_MD5(copy_filepath);
            string conv_checksum = Calculate_MD5(conv_filepath);

            // Output result in open CSV file
            var newLine1 = string.Format($"{org_filepath};{org_filesize_kb};{compare_success};{conv_filepath};{conv_filesize_kb};");
            csv.AppendLine(newLine1);

            // Close CSV file to log results
            string CSV_filepath = results_directory + "\\4_Archive_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Zip the output directory
            ZIP_Directory(results_directory);

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine("X spreadsheets failed file format validation");
            Console.WriteLine($"x out of {numTOTAL} spreadsheets were archived");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}"); //Filepath is incorrect. It is not the zipped path
            Console.WriteLine("Archiving finished");
            Console.WriteLine("---");

        }

    }

}
