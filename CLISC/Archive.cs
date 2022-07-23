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
        public void Archive(string argument0, string argument1, string argument2, string results_directory)
        {

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original checksum;Copy filepath; Copy filename; Conversion identified;Conversion filepath;Conversion filename;Conversion checksum;File format validated, Data quality message");
            csv.AppendLine(newLine0);

            // Create enumeration of converted spreadsheets based on input directory
            string docCollection = results_directory + "\\docCollection";
            List<string> docCollection_enumeration = Enumerate_docCollection(argument0, docCollection);

            foreach (var file in docCollection_enumeration)
            {

                // Create instance for finding file information
                FileInfo file_info = new FileInfo(file);

                // Combine data types to original spreadsheets
                conv_extension = file_info.Extension;
                conv_filename = file_info.Name;
                conv_filepath = file_info.FullName;

                // Rename and move converted spreadsheets

                // Copy original spreadsheets


                // Validate file format standards
                bool valid_file_format;

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
                       valid_file_format = Validate_OOXML(argument1);
                        break;
                }

                // Calculate checksums
                string copy_checksum = Calculate_MD5(org_filepath);
                string conv_checksum = Calculate_MD5(conv_filepath);

                // Check for data quality requirements
                string dataquality_message = "";

                // Output result in open CSV file
                //var newLine1 = string.Format($"{org_filepath};{org_filename};{copy_filepath};{copy_filename};{convert_success};{conv_filepath};{conv_filename};{conv_checksum};{validation_message};{dataquality_message}");
                //csv.AppendLine(newLine1);

            }

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
