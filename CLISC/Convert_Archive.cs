﻿using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;
using DocumentFormat.OpenXml.Packaging;

namespace CLISC
{
    public partial class Conversion
    {
        string? xlsx_conv_extension = null;
        string? xlsx_conv_filename = null;
        string? xlsx_conv_filepath = null;
        string? ods_conv_extension = null;
        string? ods_conv_filename = null;
        string? ods_conv_filepath = null;

        // Convert spreadsheets method
        public List<fileIndex> Convert_Spreadsheets_Archive(string inputdir, bool recurse, string Results_Directory)
        {
            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original file format;Convert filepath;Convert filename;Convert file format;Convert filepath;ODS Convert filename;ODS file format;Convert success;Convert Message");
            csv.AppendLine(newLine0);

            // Create lists
            List<orgIndex> Org_File_List = orgIndex.Org_Files(inputdir, recurse);
            List<fileIndex> File_List = new List<fileIndex>();

            // Create subdirectory (docCollection) for converted spreadsheet files
            string docCollection = Results_Directory + "\\docCollection";
            DirectoryInfo Output_Dir = Directory.CreateDirectory(docCollection);

            // Loop spreadsheets based on enumeration
            foreach (var entry in Org_File_List)
            {
                // Create data types for original files and connect to list of original files
                string org_extension = entry.Org_Extension;
                string org_filename = entry.Org_Filename;
                string org_filepath = entry.Org_Filepath;

                // Create new subdirectory for the spreadsheet
                file_folder = docCollection + "\\" + subdir_number;
                while (Directory.Exists(file_folder))
                {
                    subdir_number++;
                    file_folder = docCollection + "\\" + subdir_number;
                }
                DirectoryInfo Output_Subdir = Directory.CreateDirectory(file_folder);

                // Transform data types for copied original spreadsheet
                copy_extension = org_extension;
                copy_filename = "orgFile_" + org_filename;
                copy_filepath = file_folder + "\\" + copy_filename;

                // Copy spreadsheet 
                File.Copy(org_filepath, copy_filepath);

                // Convert spreadsheet
                try
                {
                    // Change conversion method based on file extension
                    switch (org_extension)
                    {
                        // OpenDocument file formats using LibreOffice
                        case ".fods":
                        case ".ods":
                        case ".ots":
                            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified

                            // Convert to XLSX
                            convert_success = Convert_from_OpenDocument(copy_filepath, file_folder);
                            // Because of previous bug we must rename converted spreadsheet
                            if (convert_success == true)
                            {
                                string[] filename = Directory.GetFiles(file_folder, "*.xlsx");
                                if (filename.Length > 0)
                                {
                                    // Rename converted spreadsheet
                                    string old_filename = filename[0];
                                    string new_filename = file_folder + "\\1.xlsx";
                                    File.Move(old_filename, new_filename);
                                    // Transform datatypes
                                    xlsx_conv_extension = ".xlsx";
                                    xlsx_conv_filename = "1" + xlsx_conv_extension;
                                    xlsx_conv_filepath = file_folder + "\\1.xlsx";
                                    numCOMPLETE++;
                                }
                                // And make a copy of ODS and rename it 1.ods
                                ods_conv_extension = ".ods";
                                ods_conv_filename = "1" + ods_conv_extension;
                                ods_conv_filepath = file_folder + "\\" + ods_conv_filename;
                                File.Copy(org_filepath, ods_conv_filepath);
                            }
                            // If OpenDocument spreadsheet is password protected or corrupt
                            else if (convert_success == false)
                            {

                                Console.WriteLine($"--> something is wrong");
                            }
                            break;

                        // Microsoft Excel Add-in file formats are not converted
                        case ".xla":
                        case ".xlam":
                            // No conversion
                            // Transform data types
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[5];
                            xlsx_conv_extension = null;
                            xlsx_conv_filename = null;
                            xlsx_conv_filepath = null;
                            ods_conv_extension = null;
                            ods_conv_filename = null;
                            ods_conv_filepath = null;
                            break;

                        // Legacy Microsoft Excel file formats
                        case ".xls":
                        case ".xlt":
                            // Convert to XLSX
                            xlsx_conv_extension = ".xlsx";
                            xlsx_conv_filename = "1.xlsx";
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";
                            convert_success = Convert_from_LegacyExcel(org_filepath, copy_filepath, xlsx_conv_filepath);
                            numCOMPLETE++;

                            // And convert to ODS
                            convert_success = Convert_to_OpenDocument(xlsx_conv_filepath, file_folder);
                            // Because of previous bug we must rename converted spreadsheet
                            if (convert_success == true)
                            {
                                string[] filename = Directory.GetFiles(file_folder, "*.ods");
                                if (filename.Length > 0)
                                {
                                    // Rename converted spreadsheet
                                    string old_filename = filename[0];
                                    string new_filename = file_folder + "\\1.ods";
                                    File.Move(old_filename, new_filename);
                                    // Transform datatypes
                                    ods_conv_extension = ".ods";
                                    ods_conv_filename = "1" + ods_conv_extension;
                                    ods_conv_filepath = file_folder + "\\" + ods_conv_filename;
                                }
                            }
                            break;
                        
                        // Using LibreOffice
                        case ".xlsb":
                            // Convert to XLSX
                            convert_success = Convert_from_OpenDocument(copy_filepath, file_folder);
                            // Because of previous bug we must rename converted spreadsheet
                            if (convert_success == true)
                            {
                                string[] filename = Directory.GetFiles(file_folder, "*.xlsx");
                                if (filename.Length > 0)
                                {
                                    // Rename converted spreadsheet
                                    string old_filename = filename[0];
                                    string new_filename = file_folder + "\\1.xlsx";
                                    File.Move(old_filename, new_filename);
                                    // Transform datatypes
                                    xlsx_conv_extension = ".xlsx";
                                    xlsx_conv_filename = "1" + xlsx_conv_extension;
                                    xlsx_conv_filepath = file_folder + "\\" + xlsx_conv_filename;
                                    numCOMPLETE++;
                                }

                                // And convert to ODS
                                convert_success = Convert_to_OpenDocument(xlsx_conv_filepath, file_folder);
                                // Because of previous bug we must rename converted spreadsheet
                                if (convert_success == true)
                                {
                                    filename = Directory.GetFiles(file_folder, "*.ods");
                                    if (filename.Length > 0)
                                    {
                                        // Rename converted spreadsheet
                                        string old_filename = filename[0];
                                        string new_filename = file_folder + "\\1.ods";
                                        File.Move(old_filename, new_filename);
                                        // Transform datatypes
                                        ods_conv_extension = ".ods";
                                        ods_conv_filename = "1" + ods_conv_extension;
                                        ods_conv_filepath = file_folder + "\\" + ods_conv_filename;
                                    }
                                }
                                // If OpenDocument spreadsheet is password protected or corrupt
                                else if (convert_success == false)
                                {

                                    Console.WriteLine($"--> something is wrong");
                                }
                            }
                            break;

                        case ".xlsm":
                        case ".xltm":
                        case ".xltx":
                            // Transform data types for converted spreadsheets
                            xlsx_conv_extension = ".xlsx";
                            xlsx_conv_filename = "1.xlsx";
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";

                            // Convert to XLSX
                            convert_success = Convert_to_OOXML_Transitional(copy_filepath, xlsx_conv_filepath);
                            numCOMPLETE++;

                            // And convert to ODS
                            convert_success = Convert_to_OpenDocument(xlsx_conv_filepath, file_folder);
                            // Because of previous bug we must rename converted spreadsheet
                            if (convert_success == true)
                            {
                                string[] filename = Directory.GetFiles(file_folder, "*.ods");
                                if (filename.Length > 0)
                                {
                                    // Rename converted spreadsheet
                                    string old_filename = filename[0];
                                    string new_filename = file_folder + "\\1.ods";
                                    File.Move(old_filename, new_filename);
                                    // Transform datatypes
                                    ods_conv_extension = ".ods";
                                    ods_conv_filename = "1" + ods_conv_extension;
                                    ods_conv_filepath = file_folder + "\\" + ods_conv_filename;
                                }
                            }
                            break;

                        case ".xlsx":

                                // Open to find Strict conformance
                                SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(org_filepath, false);
                                bool? strict = spreadsheet.StrictRelationshipFound;
                                spreadsheet.Close();
                                // No conversion
                                // Transform data types
                                numXLSX_noconversion++;
                                convert_success = null;
                                error_message = error_messages[6] + " - " + error_messages[9];
                                xlsx_conv_extension = ".xlsx";
                                xlsx_conv_filename = "1.xlsx";
                                xlsx_conv_filepath = file_folder + "\\1.xlsx";
                                // Copy and rename XLSX
                                File.Copy(copy_filepath, xlsx_conv_filepath);

                                // And convert to ODS
                                convert_success = Convert_to_OpenDocument(xlsx_conv_filepath, file_folder);
                                // Because of previous bug we must rename converted spreadsheet
                                if (convert_success == true)
                                {
                                    string[] filename = Directory.GetFiles(file_folder, "*.ods");
                                    if (filename.Length > 0)
                                    {
                                        // Rename converted spreadsheet
                                        string old_filename = filename[0];
                                        string new_filename = file_folder + "\\1.ods";
                                        File.Move(old_filename, new_filename);
                                        // Transform datatypes
                                        ods_conv_extension = ".ods";
                                        ods_conv_filename = "1" + ods_conv_extension;
                                        ods_conv_filepath = file_folder + "\\" + ods_conv_filename;
                                    }
                                }
                            break;
                    }
                }
                // If spreadsheet is password protected or corrupt
                catch (FileFormatException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }
                catch (InvalidDataException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }
                // If file is corrupt and cannot be opened for XML schema validation
                catch (OpenXmlPackageException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[7];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }
                // If LibreOffice is not installed
                catch (Win32Exception)
                {
                    // COde to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[3];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }
                // NPOI encryption
                catch (NPOI.Util.RecordFormatException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }
                finally
                {
                    // Inform user
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Conversion {convert_success}");
                    if (convert_success == true)
                    {
                        Console.WriteLine($"--> Conversion saved to: {xlsx_conv_filepath}");
                        Console.WriteLine($"--> Conversion saved to: {ods_conv_filepath}");
                    }
                    else if (error_message != null || error_message == error_messages[6])
                    {
                        Console.WriteLine($"--> error_message");
                    }

                    // Add copied and converted spreadsheets file info to index of files
                    File_List.Add(new fileIndex { File_Folder = file_folder, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = copy_filepath, Copy_Filename = copy_filename, Copy_Extension = copy_extension, XLSX_Conv_Filepath = xlsx_conv_filepath, XLSX_Conv_Filename = xlsx_conv_filename, XLSX_Conv_Extension = xlsx_conv_extension, ODS_Conv_Filepath = ods_conv_filepath, ODS_Conv_Filename = ods_conv_filename, ODS_Conv_Extension = ods_conv_extension, Convert_Success = convert_success });

                    // Output result in open CSV file
                    var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{xlsx_conv_filepath};{ods_conv_filename};{convert_success};{error_message}");
                    csv.AppendLine(newLine2);
                }
            }
            // Close CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\2_Convert_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());

            // Inform user of results
            Convert_Results_Archive();

            return File_List;
        }

        public static void Convert_Results_Archive()
        {
            numTOTAL_conv = numCOMPLETE + numXLSX_noconversion;

            Console.WriteLine("---");
            Console.WriteLine("CONVERT RESULTS");
            Console.WriteLine("---");
            Console.WriteLine($"{numCOMPLETE} out of {Count.numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"{numXLSX_noconversion} spreadsheets were already .xlsx");
            Console.WriteLine($"{numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {Spreadsheet.CSV_filepath}");
            Console.WriteLine("Conversion ended");
            Console.WriteLine("---");
        }
    }
}
