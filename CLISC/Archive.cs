﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;

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
            bool archive_req_accept = true;

            // Open CSV file to log archive results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;Original Checksum;Copy Filepath;Copy Checksum;Convert Exists;XLSX Convert Filepath;XLSX Checksum;XLSX File Format Validation;ODS Convert Filepath; ODS checksum; ODS file Format Validation;Archival Requirements");
            csv.AppendLine(newLine0);

            // Open CSV file to log validation results
            var csv2 = new StringBuilder();
            var newLine2_1 = string.Format($"Original Filepath;XLSX Convert Filepath;Validity;Error Number;Description;Error Type;Node;Path;Part;Related Node;Related Node Inner Text");
            csv2.AppendLine(newLine2_1);

            // Open CSV file to log archival requirements results
            var csv3 = new StringBuilder();
            var newLine3_1 = string.Format($"Original Filepath;XLSX Convert Filepath;Cell Values Exist;Data Connections Removed;External Relationships Removed;RTD Functions Removed;Embedded Objects Alert;Hyperlinks Alert");
            csv3.AppendLine(newLine3_1);

            foreach (fileIndex entry in File_List) // Loop through each file
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

                if (File.Exists(xlsx_conv_filepath))
                {
                    Console.WriteLine(org_filepath); // Inform user of original filepath
                    string folder_number = Path.GetFileName(Path.GetDirectoryName(xlsx_conv_filepath));
                    Console.WriteLine($"--> Conversion analyzed: {folder_number}\\1.xlsx"); // Inform user of analyzed filepath

                    // Convert to .xlsx Strict conformance using Excel
                    bool? strict = null;
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx_conv_filepath, false))
                    {
                        strict = spreadsheet.StrictRelationshipFound; // Identify if already Strict
                    }
                    if (strict == true)
                    {
                        Console.WriteLine("--> Spreadsheet is already Strict conformant");
                    }
                    else
                    {
                        Conversion con = new Conversion();
                        convert_success = con.Convert_Transitional_to_Strict(xlsx_conv_filepath, xlsx_conv_filepath);
                        Console.WriteLine("--> Converted to Strict conformance");
                    }

                    // Validate
                    Validation validate = new Validation();
                    List<Validation> xlsx_validation_list = validate.Validate_OOXML(org_filepath, xlsx_conv_filepath, Results_Directory);

                    foreach (Validation error in xlsx_validation_list) // Get information from validation list
                    {
                        xlsx_validity = error.Validity;
                        int? error_number = error.Error_Number;
                        string? error_description = error.Error_Description;
                        string? error_type = error.Error_Type;
                        string? error_node = error.Error_Node;
                        string? error_path = error.Error_Path;
                        string? error_part = error.Error_Part;
                        string? error_relatednode = error.Error_RelatedNode;
                        string? error_relatedtext = error.Error_RelatedNode_InnerText;

                        if (xlsx_validity == "Invalid")
                        {
                            // If invalid write to CSV validation log
                            var newLine2_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};{xlsx_validity};{error_number};{error_description};{error_type};{error_node};{error_path};{error_part};{error_relatednode};{error_relatedtext}");
                            csv2.AppendLine(newLine2_2);

                            // Reset data types, for correct CSV file output
                            error_relatednode = null;
                            error_relatedtext = null;
                        }
                    }

                    // Check .xlsx for archival requirements
                    Tuple<bool, int, int, int, int, int> pidgeon = Check_XLSX_Requirements(xlsx_conv_filepath);

                    // Receive infomration from tuple
                    bool data = pidgeon.Item1;
                    int connections = pidgeon.Item2;
                    int extrels = pidgeon.Item3;
                    int rtdfunctions = pidgeon.Item4;
                    int embedobj = pidgeon.Item5;
                    int hyperlinks = pidgeon.Item6;

                    if (data == false || embedobj > 0)
                    {
                        archive_req_accept = false;
                    }

                    // Write to CSV archival requirements log
                    var newLine3_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};{data};{connections};{extrels};{rtdfunctions};{embedobj};{hyperlinks}");
                    csv3.AppendLine(newLine3_2);

                    // Transform data according to archiving requirements
                    Transform_Requirements(xlsx_conv_filepath);

                    // Calculate checksum
                    xlsx_conv_checksum = Calculate_MD5(xlsx_conv_filepath);
                    Console.WriteLine("--> Checksum was calculated");
                }
                if (entry.ODS_Conv_Extension == ".ods")
                {
                    // Make an .ods copy
                    Conversion con = new Conversion();
                    convert_success = con.Convert_to_OpenDocument_ExcelInterop(xlsx_conv_filepath, ods_conv_filepath);
                    // Inform user
                    string folder_number = Path.GetFileName(Path.GetDirectoryName(ods_conv_filepath));
                    Console.WriteLine($"--> Conversion analyzed: {folder_number}\\1.ods");
                    Console.WriteLine("--> File format validation for .ods is not supported");
                    Console.WriteLine($"--> Archival requirements identical to {folder_number}\\1.xlsx");

                    // Calculate checksum
                    ods_conv_checksum = Calculate_MD5(ods_conv_filepath);
                    Console.WriteLine("--> Checksum was calculated");
                }

                // Calculate checksums for original and copied files
                org_checksum = Calculate_MD5(org_filepath);
                copy_checksum = Calculate_MD5(copy_filepath);

                // Output result in open CSV validation log
                var newLine1 = string.Format($"{org_filepath};{org_checksum};{copy_filepath};{copy_checksum};{convert_success};{xlsx_conv_filepath};{xlsx_conv_checksum};{xlsx_validity};{ods_conv_filepath};{ods_conv_checksum};.ods validation not supported;{archive_req_accept}");
                csv.AppendLine(newLine1);

                // Reset data types to fix bug in CSV log, if converted spreadsheet does not exist
                xlsx_conv_filepath = "";
                xlsx_conv_extension = "";
                ods_conv_filepath = "";
                ods_conv_extension = "";
                xlsx_conv_checksum = "";
                ods_conv_checksum = "";
                archive_req_accept = true;
            }

            // Close validation CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\4a_Validation_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv2.ToString());

            // Close archival requirements CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\4b_Requirements_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv3.ToString());

            // Close archive CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());

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
