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
            string copy_checksum = "";
            string conv_checksum = "";
            string dataquality_message = "";
            string validation_message = "";

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original checksum;Copy filepath; Copy filename; Conversion identified;Conversion filepath;Conversion filename;Conversion checksum;File format validated, Data quality message");
            csv.AppendLine(newLine0);

            // Validate file format standards
            Console.WriteLine("--> FILE FORMAT VALIDATION");
            Console.WriteLine("---");
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                conv_filepath = entry.Conv_Filepath;
                conv_extension = entry.Conv_Extension;

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
                }
            }
            Console.WriteLine("Validation ended");

            // Perform data quality actions
            Console.WriteLine("---");
            Console.WriteLine("--> DATA QUALITY REQUIREMENTS");
            Console.WriteLine("---");
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                conv_filepath = entry.Conv_Filepath;

                // Perform data quality actions
                dataquality_message = Transform_DataQuality(conv_filepath);
            }
            Console.WriteLine("Data quality requirements ended");

            //Calculate checksums
            Console.WriteLine("---");
            Console.WriteLine("--> CALCULATE CHECKSUMS");
            Console.WriteLine("---");
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                org_filepath = entry.Org_Filepath;
                copy_filepath = entry.Copy_Filepath;
                conv_filepath = entry.Conv_Filepath;

                // Calculate checksums
                org_checksum = Calculate_MD5(org_filepath);
                copy_checksum = Calculate_MD5(copy_filepath);
                conv_checksum = Calculate_MD5(conv_filepath);
            }
            Console.WriteLine("All file checksums were calculated");
            Console.WriteLine("Calculate checksums ended");

            // Output result in open CSV file
            var newLine1 = string.Format($"{org_filepath};{org_checksum};{copy_filepath};{copy_checksum};{conv_filepath};{conv_checksum};{validation_message};{dataquality_message}");
            csv.AppendLine(newLine1);

            // Close CSV file to log results. MUST HAPPEN BEFORE ZIP
            string CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Zip the output directory
            Console.WriteLine("---");
            Console.WriteLine("--> ZIP DIRECTORY");
            Console.WriteLine("---");
            try
            {
                ZIP_Directory(Results_Directory);
                Console.WriteLine("Zip completed");
                Console.WriteLine($"\"The zipped archive directory was saved to: \" + {Results_Directory} + \".zip\"");
                Console.WriteLine("Zip ended");
            }
            catch (SystemException)
            {
                Console.WriteLine("Zip failed");
                Console.WriteLine("Zip ended");
            }

            // Inform user of archiving results
            Console.WriteLine("---");
            Console.WriteLine("ARCHIVE RESULTS");
            Console.WriteLine("---");
            Console.WriteLine($"{valid_files} converted spreadsheets were valid");
            Console.WriteLine($"{invalid_files} converted spreadsheets were invalid");
            Console.WriteLine($"{extrels_files} converted spreadsheets had external relationships. All relationships were removed");
            Console.WriteLine($"{embedobj_files} converted spreadsheets have embedded objects. Nothing was changed");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
            Console.WriteLine("Archiving ended");
            Console.WriteLine("---");
        }
    }
}
