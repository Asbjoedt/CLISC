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

            string folder = "";
            string org_filepath = "";
            string copy_filepath = "";
            string conv_filepath = "";
            string conv_extension = "";
            string copy_extension = "";
            string org_checksum = "";
            string conv_checksum = "";
            string dataquality_message = "";
            string validation_message = "";

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original checksum;Copy filepath; Copy filename; Conversion identified;Conversion filepath;Conversion filename;Conversion checksum;File format validated, Data quality message");
            csv.AppendLine(newLine0);

            // Perform data quality actions
            Console.WriteLine("--> DATA SANITATION");
            Console.WriteLine("---");
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                folder = entry.File_Folder;
                org_filepath = entry.Org_Filepath;
                copy_filepath = entry.Copy_Filepath;
                conv_filepath = entry.Conv_Filepath;
                conv_extension = entry.Conv_Extension;
                copy_extension = entry.Copy_Extension;

                // Perform data quality requirements
                dataquality_message = Manipulate_DataQuality(conv_filepath);
            }
            // Inform user of sanitizer data results
            Console.WriteLine("---");
            Console.WriteLine("Data sanitation finished");
            Console.WriteLine("---");

            // Validate file format standards
            Console.WriteLine("--> VALIDATION");
            Console.WriteLine("---");
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                folder = entry.File_Folder;
                org_filepath = entry.Org_Filepath;
                copy_filepath = entry.Copy_Filepath;
                conv_filepath = entry.Conv_Filepath;
                conv_extension = entry.Conv_Extension;
                copy_extension = entry.Copy_Extension;

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
                        Console.WriteLine(conv_extension);
                        Console.WriteLine("--> The file format is not supported in validation workflow");
                        break;
                }
            }
            // Inform user of validation results
            Console.WriteLine("---");
            Console.WriteLine($"{valid_files} spreadsheets were valid");
            Console.WriteLine($"{invalid_files} spreadsheets were invalid");
            Console.WriteLine("Validation finished");
            Console.WriteLine("---");

            //Calculate checksums
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                folder = entry.File_Folder;
                org_filepath = entry.Org_Filepath;
                copy_filepath = entry.Copy_Filepath;
                conv_filepath = entry.Conv_Filepath;
                conv_extension = entry.Conv_Extension;
                copy_extension = entry.Copy_Extension;

                // Calculate checksums
                org_checksum = Calculate_MD5(org_filepath);
                conv_checksum = Calculate_MD5(conv_filepath);
            }

            // Output result in open CSV file
            var newLine1 = string.Format($"{org_filepath};{org_checksum};{copy_filepath};{conv_filepath};{conv_checksum};{validation_message};{dataquality_message}");
            csv.AppendLine(newLine1);

            // Close CSV file to log results. Must be before the zip
            string CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Zip the output directory
            Console.WriteLine("--> ZIP DIRECTORY");
            Console.WriteLine("---");
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
