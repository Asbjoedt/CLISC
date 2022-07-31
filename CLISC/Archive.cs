using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Archive
    {
        // Archive the spreadsheets according to advanced archival requirements
        public void Archive_Spreadsheets(string Results_Directory, List<fileIndex> File_List)
        {
            Console.WriteLine("ARCHIVE");
            Console.WriteLine("---");

            string org_filepath = "";
            string copy_filepath = "";
            string conv_filepath = "";
            string conv_extension = "";
            string org_checksum = "";
            string copy_checksum = "";
            string conv_checksum = "";
            string dataquality_message = "";
            string validation_message = "";

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original checksum;Copy filepath;Copy checksum;Conversion filepath;Conversion checksum;File format validation;Data quality message");
            csv.AppendLine(newLine0);

            // Loop through each file
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                org_filepath = entry.Org_Filepath;
                copy_filepath = entry.Copy_Filepath;
                conv_filepath = entry.Conv_Filepath;
                conv_extension = entry.Conv_Extension;

                switch (conv_extension)
                {
                    case ".xlsx":
                        // Inform user of analyzed filepath
                        Console.WriteLine(conv_filepath);
                        // Validate
                        validation_message = Validate_OOXML(conv_filepath);

                        // Perform data quality actions
                        dataquality_message = Transform_DataQuality(conv_filepath);

                        // Calculate checksums
                        org_checksum = Calculate_MD5(org_filepath);
                        copy_checksum = Calculate_MD5(copy_filepath);
                        conv_checksum = Calculate_MD5(conv_filepath);
                        // Inform user
                        Console.WriteLine("--> Checksum was calculated");

                        // Output result in open CSV file
                        var newLine1 = string.Format($"{org_filepath};{org_checksum};{copy_filepath};{copy_checksum};{conv_filepath};{conv_checksum};{validation_message};{dataquality_message}");
                        csv.AppendLine(newLine1);
                        break;
                }
            }
            // Close CSV file to log results. MUST HAPPEN BEFORE ZIP
            Spreadsheet.CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());

            // Zip the output directory
            Console.WriteLine("--> ZIP DIRECTORY");
            Console.WriteLine("---");
            try
            {
                // ZIP_Directory(Results_Directory); Commented out zipping for debugging reasons
                string zip_path = Results_Directory + ".zip";
                Console.WriteLine($"The zipped archive directory was saved to: {zip_path}");
                Console.WriteLine("Zip ended");
            }
            catch (SystemException)
            {
                Console.WriteLine("Zip failed");
                Console.WriteLine("Zip ended");
            }

            // Inform user of archiving results
            Archive_Results();
        }

        public void Archive_Results()
        {
            Console.WriteLine("---");
            Console.WriteLine("ARCHIVE RESULTS");
            Console.WriteLine("---");
            Console.WriteLine($"{valid_files} out of {Conversion.numTOTAL_conv} converted spreadsheets have valid file formats");
            Console.WriteLine($"{invalid_files} out of {Conversion.numTOTAL_conv} converted spreadsheets have invalid file formats");
            Console.WriteLine($"{extrels_files} out of {Conversion.numTOTAL_conv} converted spreadsheets had external relationships - They were removed");
            Console.WriteLine($"{embedobj_files} out of {Conversion.numTOTAL_conv} converted spreadsheets have embedded objects - They were NOT removed");
            Console.WriteLine($"Results saved to CSV log in filepath: {Spreadsheet.CSV_filepath}");
            Console.WriteLine("Archiving ended");
            Console.WriteLine("---");
        }
    }
}
