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
        public void Archive(string argument0, string argument1, string argument2, string Results_Directory, List<fileIndex> File_List)
        {
            Console.WriteLine("ARCHIVE");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original checksum;Copy filepath; Copy filename; Conversion identified;Conversion filepath;Conversion filename;Conversion checksum;File format validated, Data quality message");
            csv.AppendLine(newLine0);

            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                string org_filepath = entry.Org_Filepath;
                string conv_filepath = entry.Conv_Filepath;
                string folder = entry.File_Folder;
                string conv_extension = entry.Conv_Extension;
                string copy_extension = entry.Copy_Extension;

                // Validate file format standards
                Console.WriteLine("--> VALIDATION");
                Console.WriteLine("---");

                switch (conv_extension)
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
                        validation_message = Validate_OOXML(conv_filepath);
                        break;

                    default:
                        Console.WriteLine(copy_extension);
                        Console.WriteLine("--> The file format is not supported in validation routine");
                        break;
                }
                // Inform user of validation results
                Console.WriteLine("---");
                Console.WriteLine($"{valid_files} spreadsheets were valid");
                Console.WriteLine($"{invalid_files} spreadsheets were invalid");
                Console.WriteLine("Validation finished");
                Console.WriteLine("---");

                // Calculate checksums
                string copy_checksum = Calculate_MD5(org_filepath);
                string conv_checksum = Calculate_MD5(conv_filepath);

                // Check for data quality requirements

                // Validate file format standards
                Console.WriteLine("--> DATA QUALITY");
                Console.WriteLine("---");

                string dataquality_message = "";

                Console.WriteLine("Data quality finished");
                Console.WriteLine("---");

                // Output result in open CSV file
                //var newLine1 = string.Format($"{org_filepath};{org_filename};{copy_filepath};{copy_filename};{convert_success};{conv_filepath};{conv_filename};{conv_checksum};{validation_message};{dataquality_message}");
                //csv.AppendLine(newLine1);
            }

            // Close CSV file to log results. Must be before the zip
            string CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Zip the output directory
            Console.WriteLine("--> ZIP DIRECTORY");
            Console.WriteLine("---");
            bool zip = true;
            try
            {
            ZIP_Directory(Results_Directory);
                Console.WriteLine($"Zipped output archive directory saved at: {argument1}");
                Console.WriteLine("Zip finished");
            }
            catch (SystemException)
            {
                Console.WriteLine("Zip failed");
            }

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine($"{invalid_files} spreadsheets failed file format validation");
            Console.WriteLine($"x out of {numTOTAL} spreadsheets were archived");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
            Console.WriteLine("Archiving finished");
            Console.WriteLine("---");
        }
    }
}
