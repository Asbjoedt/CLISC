﻿using System;
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
        public static int valid_files = 0;
        public static int invalid_files = 0;
        public static int cellvalue_files = 0;
        public static int connections_files = 0;
        public static int cellreferences_files = 0;
        public static int rtdfunctions_files = 0;
        public static int printersettings_files = 0;
        public static int extobj_files = 0;
        public static int embedobj_files = 0;
        public static int hyperlinks_files = 0;
        public static int activesheet_files = 0;
        public static int absolutepath_files = 0;
        public static int vbaproject_files = 0;
        public static int metadata_files = 0;

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
            bool archive_req_accept = false;
            bool data = false;
            int connections = 0;
            int cellreferences = 0;
            int rtdfunctions = 0;
            int printersettings = 0;
            int extobj = 0;
            int embedobj = 0;
            int hyperlinks = 0;
            bool activesheet = false;
            bool absolutepath = false;
            bool vbaprojects = false;
            bool metadata = false;

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
            var newLine3_1 = string.Format($"Original Filepath;XLSX Convert Filepath;Cell Values Exist;Data Connections Removed;Cell References Removed;RTD Functions Removed;Printersettings Removed;External Objects Removed;Embedded Objects Alert;Hyperlinks Alert;Active Sheet changed");
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
                        // Inform user of archival requirements
                        Console.WriteLine(org_filepath); // Inform user of original filepath
                        string folder_number = Path.GetFileName(Path.GetDirectoryName(xlsx_conv_filepath));
                        Console.WriteLine($"--> Analyzing: {folder_number}\\1.xlsx"); // Inform user of analyzed filepath

                        // Convert to .xlsx Strict conformance
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
                            convert_success = con.Convert_Transitional_to_Strict_ExcelInterop(xlsx_conv_filepath, xlsx_conv_filepath);
                            //con.Convert_Transitional_to_Strict(xlsx_conv_filepath);
                            if (convert_success == true)
                            {
                                Console.WriteLine("--> Converted to Strict conformance");
                            }
                            else
                            {
                                Console.WriteLine("--> Failed to convert to Strict conformance");
                            }
                            
                        }

                        // Check .xlsx for archival requirements
                        Archive_Requirements arc = new Archive_Requirements();
                        List<Archive_Requirements> pidgeon = arc.Check_XLSX_Requirements(xlsx_conv_filepath);
                        foreach (var item in pidgeon)
                        {
                            // Receive information
                            data = item.Data;
                            connections = item.Connections;
                            cellreferences = item.CellReferences;
                            rtdfunctions = item.RTDFunctions;
                            printersettings = item.PrinterSettings;
                            extobj = item.ExternalObj;
                            embedobj = item.EmbedObj;
                            hyperlinks = item.Hyperlinks;
                            activesheet = item.ActiveSheet;
                            absolutepath = item.AbsolutePath;
                            vbaprojects = item.VBAProjects;
                            metadata = item.Metadata;
                        }

                        // Transform data according to archiving requirements
                        if (connections > 0)
                        {
                            connections_files++;
                            arc.Remove_DataConnections(xlsx_conv_filepath);
                        }
                        if (cellreferences > 0)
                        {
                            cellreferences_files++;
                            arc.Remove_CellReferences(xlsx_conv_filepath);
                        }
                        if (rtdfunctions > 0)
                        {
                            rtdfunctions_files++;
                            arc.Remove_RTDFunctions(xlsx_conv_filepath);
                        }
                        if (printersettings > 0)
                        {
                            printersettings_files++;
                            arc.Remove_PrinterSettings(xlsx_conv_filepath);
                        }
                        if (extobj > 0)
                        {
                            extobj_files++;
                            arc.Remove_ExternalObjects(xlsx_conv_filepath);
                        }
                        if (embedobj > 0)
                        {
                            embedobj_files++;
                            arc.Remove_EmbeddedObjects(xlsx_conv_filepath);
                        }
                        if (hyperlinks > 0)
                        {
                            hyperlinks_files++;
                        }
                        if (activesheet == true)
                        {
                            activesheet_files++;
                            arc.Activate_FirstSheet(xlsx_conv_filepath);
                        }
                        if (absolutepath == true)
                        {
                            absolutepath_files++;
                            arc.Remove_AbsolutePath(xlsx_conv_filepath);
                        }
                        if (data == false)
                        {
                            cellvalue_files++;
                            archive_req_accept = true;
                        }
                        if (vbaprojects == true)
                        {
                            vbaproject_files++;
                            arc.Remove_VBA(xlsx_conv_filepath);
                        }
                        if (metadata == true)
                        {
                            metadata_files++;
                            //arc.Remove_Metadata(xlsx_conv_filepath);
                        }

                        // Write to CSV archival requirements log
                        var newLine3_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};{data};{connections};{cellreferences};{rtdfunctions};{printersettings};{extobj};{embedobj};{hyperlinks};{activesheet}");
                        csv3.AppendLine(newLine3_2);

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
                        // Write to CSV archival requirements log
                        var newLine3_2 = string.Format($"{org_filepath};{xlsx_conv_filepath};;;;;;");
                        csv3.AppendLine(newLine3_2);

                        xlsx_validity = "Invalid";
                        archive_req_accept = false;
                    }

                    // Calculate checksum
                    xlsx_conv_checksum = Calculate_MD5(xlsx_conv_filepath);
                    Console.WriteLine("--> Checksum (MD5) was calculated");
                }
                if (entry.ODS_Conv_Extension == ".ods")
                {
                    // Make an .ods copy
                    Conversion con = new Conversion();
                    convert_success = con.Convert_to_OpenDocument(xlsx_conv_filepath, file_folder);

                    // Inform user
                    string folder_number = Path.GetFileName(Path.GetDirectoryName(ods_conv_filepath));
                    Console.WriteLine($"--> File saved to: {folder_number}\\1.ods");
                    Console.WriteLine($"--> Analyzing: {folder_number}\\1.ods");
                    Console.WriteLine($"--> Archival requirements identical to {folder_number}\\1.xlsx");

                    // Validate .ods
                    Validation val = new Validation();
                    //ods_validity = val.Validate_OpenDocument(ods_conv_filepath);

                    // Calculate checksum
                    ods_conv_checksum = Calculate_MD5(ods_conv_filepath);
                    Console.WriteLine("--> Checksum (MD5) was calculated");
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
