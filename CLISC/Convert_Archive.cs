using System;
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
        string org_extension = "";
        string org_filename = "";
        string org_filepath = "";
        string? xlsx_conv_extension = null;
        string? xlsx_conv_filename = null;
        string? xlsx_conv_filepath = null;
        string? ods_conv_extension = null;
        string? ods_conv_filename = null;
        string? ods_conv_filepath = null;

        // Convert spreadsheets method
        public List<fileIndex> Convert_Spreadsheets_Archive(string function, string inputdir, bool recurse, string Results_Directory)
        {
            Console.WriteLine("---");
            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;Original Filename;Original Fileformat;XLSX Convert Filepath;ODS Convert Filepath;Convert Success;Convert Message");
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
                org_extension = entry.Org_Extension;
                org_filename = entry.Org_Filename;
                org_filepath = entry.Org_Filepath;

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
                        case ".gsheet":
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[11];
                            break;

                        case ".numbers":
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[12];
                            break;

                        // OpenDocument file formats
                        case ".fods":
                        case ".ods":
                        case ".ots":
                            // Convert to XLSX Transitional
                            convert_success = Convert_from_OpenDocument(function, copy_filepath, file_folder); // Using LibreOffice

                            break;

                        // Microsoft Excel Add-in file formats are not converted
                        case ".xla":
                        case ".xlam":
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
                            // Transform data types for converted spreadsheets
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";
                            // Convert to .xlsx Transitional using Excel Interop
                            convert_success = Convert_Legacy_ExcelInterop(copy_filepath, xlsx_conv_filepath);
                            break;

                        case ".xlsb":
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";
                            // Convert to .xlsx Transitional using Excel Interop
                            convert_success = Convert_Legacy_ExcelInterop(copy_filepath, xlsx_conv_filepath); 
                            break;

                        case ".xlsm":
                        case ".xlsx":
                        case ".xltm":
                        case ".xltx":
                            // Transform data types for converted spreadsheets
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";
                            // Convert to XLSX Transitional
                            convert_success = Convert_to_OOXML_Transitional(copy_filepath, xlsx_conv_filepath);
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
                    // Code to execute
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
 
                finally
                {
                    // Inform user
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Conversion: {convert_success}");
                    if (convert_success == false)
                    {
                        Console.WriteLine($"--> {error_message}");
                    }

                    if (convert_success == true)
                    {
                        // Transform data types
                        numCOMPLETE++;
                        xlsx_conv_extension = ".xlsx";
                        xlsx_conv_filename = "1.xlsx";
                        xlsx_conv_filepath = file_folder + "\\1.xlsx";

                        // Check for original extension already .xlsx
                        if (copy_extension == ".xlsx")
                        {
                            error_message = error_messages[6];
                        }
                        else
                        {
                            error_message = "";
                        }

                        // Make a copy in ODS
                        convert_success = Convert_to_OpenDocument(function, xlsx_conv_filepath, file_folder);
                        ods_conv_extension = ".ods";
                        ods_conv_filename = "1" + ods_conv_extension;
                        ods_conv_filepath = file_folder + "\\" + ods_conv_filename;
                        // To correct for bug, where LibreOffice overwrites the copied original of an .ods spreadsheet
                        if (!File.Exists(copy_filepath))
                        {
                            File.Copy(org_filepath, copy_filepath);
                        }

                        // Inform user
                        Console.WriteLine($"--> File saved to: {xlsx_conv_filepath}");
                        Console.WriteLine($"--> File saved to: {ods_conv_filepath}");

                    }
                    else
                    {
                        convert_success = false;
                        error_message = "Spreadsheet is password protected or corrupt";
                        xlsx_conv_extension = null;
                        xlsx_conv_filename = null;
                        xlsx_conv_filepath = null;
                        ods_conv_extension = null;
                        ods_conv_filename = null;
                        ods_conv_filepath = null;
                    }

                    // Add copied and converted spreadsheets file info to index of files
                    File_List.Add(new fileIndex { File_Folder = file_folder, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = copy_filepath, Copy_Filename = copy_filename, Copy_Extension = copy_extension, XLSX_Conv_Filepath = xlsx_conv_filepath, XLSX_Conv_Filename = xlsx_conv_filename, XLSX_Conv_Extension = xlsx_conv_extension, ODS_Conv_Filepath = ods_conv_filepath, ODS_Conv_Filename = ods_conv_filename, ODS_Conv_Extension = ods_conv_extension, Convert_Success = convert_success });

                    // Output result in open CSV file
                    var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{xlsx_conv_filepath};{ods_conv_filepath};{convert_success};{error_message}");
                    csv.AppendLine(newLine2);
                }
            }
            // Close CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\2_Convert_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());

            return File_List;
        }
    }
}
