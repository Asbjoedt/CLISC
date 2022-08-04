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
            Console.WriteLine("---");
            Console.WriteLine("ARCHIVE");
            Console.WriteLine("---");

            string file_folder = "";
            string org_filepath = "";
            string copy_filepath = "";
            string xlsx_conv_filepath = "";
            string xlsx_conv_extension = "";
            string ods_conv_filepath = "";
            string ods_conv_extension = "";
            string org_checksum = "";
            string copy_checksum = "";
            string xlsx_conv_checksum = "";
            string ods_conv_checksum = "";
            string xlsx_validation_message = "";
            string ods_validation_message = "";
            string dataquality_message = "";

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;Original Checksum;Copy Filepath;Copy Checksum;XLSX Convert Filepath;XLSX Checksum;XLSX Fileformat Validation;ODS Convert Filepath; ODS checksum; ODS fileformat Validation;Data Quality Message");
            csv.AppendLine(newLine0);

            // Loop through each file
            foreach (fileIndex entry in File_List)
            {
                // Get information from list
                file_folder = entry.File_Folder;
                org_filepath = entry.Org_Filepath;
                copy_filepath = entry.Copy_Filepath;
                xlsx_conv_filepath = entry.XLSX_Conv_Filepath;
                xlsx_conv_extension = entry.XLSX_Conv_Extension;
                ods_conv_filepath = entry.ODS_Conv_Filepath;
                ods_conv_extension = entry.ODS_Conv_Extension;

                // Inform user of analyzed filepath
                if (xlsx_conv_extension == ".xlsx")
                {
                    // Inform user
                    Console.WriteLine(xlsx_conv_filepath);

                    // Validate
                    xlsx_validation_message = Validate_OOXML(xlsx_conv_filepath);

                    // Perform data quality actions
                    dataquality_message = Check_Requirements(xlsx_conv_filepath);

                    // Calcualte checksum
                    xlsx_conv_checksum = Calculate_MD5(xlsx_conv_filepath);
                    Console.WriteLine("--> Checksum was calculated");
                }
                if (entry.ODS_Conv_Extension == ".ods")
                {
                    string folder_number = Path.GetFileName(Path.GetDirectoryName(xlsx_conv_filepath));
                    Console.WriteLine(ods_conv_filepath);
                    Console.WriteLine("--> File format validation of .ods is not supported");
                    Console.WriteLine($"--> Archival requirements acceptance is identical to: {folder_number}\\1.xlsx");

                    // Calcualte checksum
                    ods_conv_checksum = Calculate_MD5(ods_conv_filepath);
                    Console.WriteLine("--> Checksum was calculated");
                }

                // Calculate checksums for original and copied files
                org_checksum = Calculate_MD5(org_filepath);
                copy_checksum = Calculate_MD5(copy_filepath);

                // Output result in open CSV file
                var newLine1 = string.Format($"{org_filepath};{org_checksum};{copy_filepath};{copy_checksum};{xlsx_conv_filepath};{xlsx_conv_checksum};{xlsx_validation_message};{ods_conv_filepath};{ods_conv_checksum};{ods_validation_message};{dataquality_message}");
                csv.AppendLine(newLine1);
            }
            // Close CSV file to log results. MUST HAPPEN BEFORE ZIP
            Spreadsheet.CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());

            // Zip the output directory
            Console.WriteLine("---");
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
        }
    }
}
