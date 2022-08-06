using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Archive
    {
        public static int valid_files = 0;
        public static int invalid_files = 0;

        // Archive the spreadsheets according to advanced archival requirements
        public void Archive_Spreadsheets(string Results_Directory, List<fileIndex> File_List)
        {
            Console.WriteLine("---");
            Console.WriteLine("ARCHIVE");
            Console.WriteLine("---");

            string file_folder = "";
            string org_filepath = "";
            string copy_filepath = "";
            bool? convert_success;
            string xlsx_conv_filepath = "";
            string xlsx_conv_extension = "";
            string ods_conv_filepath = "";
            string ods_conv_extension = "";
            string org_checksum = "";
            string copy_checksum = "";
            string xlsx_conv_checksum = "";
            string xlsx_validity = "";
            string ods_conv_checksum = "";
            string ods_validation_message = "";
            string dataquality_message = "";

            // Open CSV file to log archive results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;Original Checksum;Copy Filepath;Copy Checksum;Convert Exists;XLSX Convert Filepath;XLSX Checksum;XLSX File Format Validation;ODS Convert Filepath; ODS checksum; ODS fileformat Validation;Data Quality Message");
            csv.AppendLine(newLine0);

            // Open CSV file to log validation results
            var csv2 = new StringBuilder();
            var newLine2_1 = string.Format($"Original Filepath;XLSX Convert Filepath;Validity;Error Number;Description;Error Type;Node;Path;Part;Related Node;Related Node Inner Text");
            csv2.AppendLine(newLine2_1);

            // Loop through each file
            foreach (fileIndex entry in File_List)
            {
                // Get information from File list
                file_folder = entry.File_Folder;
                org_filepath = entry.Org_Filepath;
                copy_filepath = entry.Copy_Filepath;
                convert_success = entry.Convert_Success;
                xlsx_conv_filepath = entry.XLSX_Conv_Filepath;
                xlsx_conv_extension = entry.XLSX_Conv_Extension;
                ods_conv_filepath = entry.ODS_Conv_Filepath;
                ods_conv_extension = entry.ODS_Conv_Extension;

                // Inform user of analyzed filepath
                if (File.Exists(xlsx_conv_filepath))
                {
                    // Inform user
                    Console.WriteLine(xlsx_conv_filepath);

                    // Validate
                    Validation validate = new Validation();

                    List<Validation> xlsx_validation_list = validate.Validate_OOXML(org_filepath, xlsx_conv_filepath, Results_Directory);

                    // Get information from validation list
                    foreach (Validation entry2 in xlsx_validation_list)
                    {
                        xlsx_validity = entry2.Validity;
                        int? error_number = entry2.Error_Number;
                        string? error_description = entry2.Error_Description;
                        string? error_type = entry2.Error_Type;
                        string? error_node = entry2.Error_Node;
                        string? error_path = entry2.Error_Path;
                        string? error_part = entry2.Error_Part;
                        string? error_relatednode = entry2.Error_RelatedNode;
                        string? error_relatedtext = entry2.Error_RelatedNode_InnerText;

                        if (xlsx_validity == "Valid")
                        {
                            valid_files++; // Add to count of valid files
                        }
                        else 
                        {
                            invalid_files++; // Add to count of invalid files

                            // If invalid write to CSV validation log
                            var newLine2_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};{xlsx_validity};{error_number};{error_description};{error_type};{error_node};{error_path};{error_part};{error_relatednode};{error_relatedtext}");
                            csv2.AppendLine(newLine2_2);

                            // Reset data types, for correctCSV file output
                            error_relatednode = null;
                            error_relatedtext = null;
                        }
                    }

                    // Perform data quality actions
                    dataquality_message = Check_Requirements(xlsx_conv_filepath);

                    // Calculate checksum
                    xlsx_conv_checksum = Calculate_MD5(xlsx_conv_filepath);
                    Console.WriteLine("--> Checksum was calculated");
                }
                if (entry.ODS_Conv_Extension == ".ods" && File.Exists(ods_conv_filepath))
                {
                    string folder_number = Path.GetFileName(Path.GetDirectoryName(xlsx_conv_filepath));
                    Console.WriteLine(ods_conv_filepath);
                    Console.WriteLine("--> File format validation for .ods is not supported");
                    Console.WriteLine($"--> Archival requirements acceptance is identical to: {folder_number}\\1.xlsx");

                    // Calculate checksum
                    ods_conv_checksum = Calculate_MD5(ods_conv_filepath);
                    Console.WriteLine("--> Checksum was calculated");
                }

                // Calculate checksums for original and copied files
                org_checksum = Calculate_MD5(org_filepath);
                copy_checksum = Calculate_MD5(copy_filepath);

                // Output result in open CSV validation log
                var newLine1 = string.Format($"{org_filepath};{org_checksum};{copy_filepath};{copy_checksum};{convert_success};{xlsx_conv_filepath};{xlsx_conv_checksum};{xlsx_validity};{ods_conv_filepath};{ods_conv_checksum};{ods_validation_message};{dataquality_message}");
                csv.AppendLine(newLine1);

                // Reset data types to fix bug in CSV log, if converted spreadsheet does not exist
                xlsx_conv_filepath = "";
                xlsx_conv_extension = "";
                ods_conv_filepath = "";
                ods_conv_extension = "";
                xlsx_conv_checksum = "";
                ods_conv_checksum = "";
                dataquality_message = "";
            }
            // Close CSV file to archive log results. MUST HAPPEN BEFORE ZIP
            Spreadsheet.CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());

            // Close CSV file to log validation results.
            Spreadsheet.CSV_filepath = Results_Directory + "\\4a_Validation_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv2.ToString());

            // Zip the output directory
            Console.WriteLine("---");
            Console.WriteLine("--> ZIP DIRECTORY");
            Console.WriteLine("---");
            try
            {
                ZIP_Directory(Results_Directory);
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
