using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Packaging;

namespace CLISC
{
    public partial class Archive
    {
        public static int cellvalue_files = 0;
        public static int metadata_files = 0;
        public static int conformance_files = 0;
        public static int connections_files = 0;
        public static int cellreferences_files = 0;
        public static int rtdfunctions_files = 0;
        public static int printersettings_files = 0;
        public static int extobj_files = 0;
        public static int activesheet_files = 0;
        public static int absolutepath_files = 0;
        public static int embedobj_files = 0;
        public static int hyperlinks_files = 0;
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
            int xlsx_errors_count = 0;
            bool? ods_validity = null;
            string ods_conv_checksum = "";
            bool archive_req_accept = true;

            // Open CSV file to log archive results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;Original Checksum;Copy Filepath;Copy Checksum;XLSX Convert Filepath;XLSX Checksum (MD5);XLSX File Format Validation;Validation Errors;ODS Convert Filepath; ODS checksum (MD5); ODS file Format Validation;Archival Requirements");
            csv.AppendLine(newLine0);

            // Open CSV file to log validation results
            var csv2 = new StringBuilder();
            var newLine2_1 = string.Format($"Original Filepath;XLSX Convert Filepath;Validity;Error Number;Id;Description;Type;Node;Path;Part;Related Node;Related Node Inner Text");
            csv2.AppendLine(newLine2_1);

            // Open CSV file to log archival requirements results
            var csv3 = new StringBuilder();
            var newLine3_1 = string.Format($"Original Filepath;XLSX Convert Filepath;Cell Values;Conformance;Data Connections;External Cell References;RTD Functions;Printersettings;External Objects;Active Sheet;Absolute Path;Embedded Objects");
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
                    try
                    {
                        // Inform user
                        Console.WriteLine(org_filepath); // Inform user of original filepath
                        string folder_number = Path.GetFileName(Path.GetDirectoryName(xlsx_conv_filepath));
                        Console.WriteLine($"--> Analyzing: {folder_number}\\1.xlsx"); // Inform user of analyzed filepath

                        // Check .xlsx for archival requirements
                        Archive_Requirements arc = new Archive_Requirements();
                        List<Archive_Requirements> arcReq = arc.Check_XLSX_Requirements(xlsx_conv_filepath);

                        // Change .xlsx according to archival requirements
                        arc.Change_XLSX_Requirements(arcReq, xlsx_conv_filepath);

                        foreach (var item in arcReq)
                        {
                            // Register and count occurences of detected breaches of archival requirements
                            if (item.Data == true)
                            {
                                cellvalue_files++;
                                archive_req_accept = false;
                            }
                            if (item.Metadata == true)
                            {
                                metadata_files++;
                            }
                            if (item.Conformance == true)
                            {
                                conformance_files++;
                            }
                            if (item.Connections > 0)
                            {
                                connections_files++;
                            }
                            if (item.CellReferences > 0)
                            {
                                cellreferences_files++;
                            }
                            if (item.RTDFunctions > 0)
                            {
                                rtdfunctions_files++;
                            }
                            if (item.PrinterSettings > 0)
                            {
                                printersettings_files++;
                            }
                            if (item.ExternalObj > 0)
                            {
                                extobj_files++;
                            }
                            if (item.ActiveSheet == true)
                            {
                                activesheet_files++;
                            }
                            if (item.AbsolutePath == true)
                            {
                                absolutepath_files++;
                            }
                            if (item.EmbedObj > 0)
                            {
                                embedobj_files++;
                            }
                            if (item.Hyperlinks > 0)
                            {
                                hyperlinks_files++;
                            }

                            // Write information to CSV archival requirements log
                            var newLine3_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};{item.Data};{item.Conformance};{item.Connections};{item.CellReferences};{item.RTDFunctions};{item.PrinterSettings};{item.ExternalObj};{item.ActiveSheet};{item.AbsolutePath};{item.EmbedObj}");
                            csv3.AppendLine(newLine3_2);
                        }

                        // Validate
                        Validation validate = new Validation();
                        if (File.Exists(xlsx_conv_filepath))
                        {
                            List<Validation> xlsx_validation_list = validate.Validate_OOXML_Hack(org_filepath, xlsx_conv_filepath, Results_Directory);

                            xlsx_errors_count = xlsx_validation_list.Count;

                            foreach (Validation info in xlsx_validation_list) // Get information from validation list
                            {
                                xlsx_validity = info.Validity;
                                int? error_number = info.Error_Number;
                                string? error_id = info.Error_Id;
                                string? error_description = info.Error_Description;
                                string? error_type = info.Error_Type;
                                string? error_node = info.Error_Node;
                                string? error_path = info.Error_Path;
                                string? error_part = info.Error_Part;
                                string? error_relatednode = info.Error_RelatedNode;
                                string? error_relatedtext = info.Error_RelatedNode_InnerText;

                                if (xlsx_validity == "Invalid")
                                {
                                    // If invalid write to CSV validation log
                                    var newLine2_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};{xlsx_validity};{error_number};{error_id};{error_description};{error_type};{error_node};{error_path};{error_part};{error_relatednode};{error_relatedtext}");
                                    csv2.AppendLine(newLine2_2);

                                    // Reset data types, for correct CSV file output
                                    error_relatednode = null;
                                    error_relatedtext = null;
                                }
                            }
                        }
                    }
                    // If spreadsheet is malformed
                    catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException)
                    {
                        xlsx_validity = "Invalid";
                        archive_req_accept = false;

                        // Write to CSV archival requirements log
                        var newLine3_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};{xlsx_validity};;;;;");
                        csv3.AppendLine(newLine3_2);
                    }

                    // Calculate checksum
                    xlsx_conv_checksum = Calculate_MD5(xlsx_conv_filepath);
                    Console.WriteLine("--> Calculate: MD5 checksum was calculated");
                }
                if (entry.ODS_Conv_Extension == ".ods")
                {
                    // Inform user of .ods operation
                    string folder_number = Path.GetFileName(Path.GetDirectoryName(ods_conv_filepath));
                    Console.WriteLine($"--> Copy saved to: {folder_number}\\1.ods");
                    Console.WriteLine($"--> Analyzing: {folder_number}\\1.ods");
                    Console.WriteLine($"--> Archival requirements identical to {folder_number}\\1.xlsx");

                    // Make an .ods copy
                    Conversion con = new Conversion();
                    convert_success = con.Convert_to_OpenDocument(xlsx_conv_filepath, file_folder);

                    // Validate .ods
                    Validation val = new Validation();
                    //ods_validity = val.Validate_OpenDocument(ods_conv_filepath);

                    // Calculate checksum
                    ods_conv_checksum = Calculate_MD5(ods_conv_filepath);
                    Console.WriteLine("--> Calculate: MD5 checksum was calculated");
                }

                // Calculate checksums for original and copied files
                org_checksum = Calculate_MD5(org_filepath);
                copy_checksum = Calculate_MD5(copy_filepath);

                // Output result in open CSV archive results log
                var newLine1 = string.Format($"{org_filepath};{org_checksum};{copy_filepath};{copy_checksum};{xlsx_conv_filepath};{xlsx_conv_checksum};{xlsx_validity};{xlsx_errors_count};{ods_conv_filepath};{ods_conv_checksum};{ods_validity};{archive_req_accept}");
                csv.AppendLine(newLine1);

                // Reset data types to fix bug in CSV log, if converted spreadsheet does not exist
                xlsx_conv_filepath = "";
                xlsx_conv_extension = "";
                ods_conv_filepath = "";
                ods_conv_extension = "";
                xlsx_conv_checksum = "";
                ods_conv_checksum = "";
                xlsx_validity = "";
                archive_req_accept = true;
            }

            // Close validation CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\4a_StandardValidation_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv2.ToString());

            // Close archival requirements CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\4b_RequirementsValidation_Results.csv";
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
