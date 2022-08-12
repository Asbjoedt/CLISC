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
using DocumentFormat.OpenXml.Spreadsheet;


namespace CLISC
{
    public partial class Conversion
    {
        // Define data types
        public static int numCOMPLETE = 0;
        public static int numFAILED = 0;
        public static bool? convert_success = null;
        static string? file_folder = null;
        static int subdir_number = 1;
        static int copy_file_number = 1;
        static int conv_file_number = 1;
        string org_extension = "";
        string org_filename = "";
        string org_filepath = "";
        static string copy_extension = "";
        static string copy_filename = "";
        static string copy_filepath = "";
        static string? conv_extension = null;
        static string? conv_filename = null;
        static string? conv_filepath = null;
        string? xlsx_conv_extension = null;
        string? xlsx_conv_filename = null;
        string? xlsx_conv_filepath = null;
        string? ods_conv_extension = null;
        string? ods_conv_filename = null;
        string? ods_conv_filepath = null;
        static string? error_message = null;
        static string[] error_messages = { "", "Legacy Excel file formats are not supported", "Binary .xlsb file format needs Excel installed with .NET programming", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet cannot be read", "Microsoft Excel Add-In file format cannot contain any cell values and is not converted", "Spreadsheet is already .xlsx file format", "Spreadsheet cannot be opened, because the XML structure is malformed", "Spreadsheet was converted to OOXML Transitional conformance", ".xlsx Strict conformance identified", "Cannot convert automatically because of irregular content", "Google Sheets are stored in the cloud and cannot be converted locally", "Apple Numbers file format is not supported", "Converted to Strict conformance", "Conversion to Strict conformance failed.", "Conversion of file has exceeded 5 min. Handle file manually" };

        // Convert spreadsheets method
        public List<fileIndex> Convert_Spreadsheets(string function, string inputdir, bool recurse, string Results_Directory)
        {
            Console.WriteLine("---");
            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original Filepath;Original Filename;Original File Format;XLSX Convert Filepath;ODS Convert Filepath;Convert Success;Convert Message");
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
                    if (function == "count&convert&compare&archive")
                    {
                        conv_filepath = file_folder + "\\1.xlsx";
                    }
                    else
                    {
                        conv_filepath = file_folder + "\\orgFile_" + Path.GetFileNameWithoutExtension(org_filename) + ".xlsx";
                    }

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
                            // Convert to XLSX Transitional using LibreOffice
                            convert_success = Convert_from_OpenDocument(function, copy_filepath, file_folder);

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
                            convert_success = Convert_Legacy_ExcelInterop(copy_filepath, conv_filepath);
                            break;

                        case ".xlsb":
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";
                            // Convert to .xlsx Transitional using Excel Interop
                            convert_success = Convert_Legacy_ExcelInterop(copy_filepath, conv_filepath);
                            break;

                        case ".xlsm":
                        case ".xlsx":
                        case ".xltm":
                        case ".xltx":
                            // Transform data types for converted spreadsheets
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";
                            // Convert to .xlsx Transitional using Open XML SDK
                            convert_success = Convert_to_OOXML_Transitional(copy_filepath, conv_filepath);
                            break;
                    }
                }
                // If spreadsheet is password protected or corrupt
                catch (FileFormatException)
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                }
                catch (InvalidDataException)
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                }
                // If file is corrupt and cannot be opened for XML schema validation
                catch (OpenXmlPackageException)
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[7];
                }
                // If .LibreOffice is not installed
                catch (Win32Exception)
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                }
                // If files used by Excel Interop are password protected or corrupt
                catch (System.Runtime.InteropServices.COMException)
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                }
                // If file conversion exceeds 5 min
                catch (TimeoutException)
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[13];
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
                        if (function == "count&convert&compare&archive")
                        {
                            // Transform data types
                            numCOMPLETE++;
                            xlsx_conv_extension = ".xlsx";
                            xlsx_conv_filename = "1.xlsx";
                            xlsx_conv_filepath = file_folder + "\\1.xlsx";
                            ods_conv_extension = ".ods";
                            ods_conv_filename = "1.ods";
                            ods_conv_filepath = file_folder + "\\1.ods";
                            error_message = "";

                            // Inform user
                            Console.WriteLine($"--> File saved to: {xlsx_conv_filepath}");
                            Console.WriteLine($"--> File saved to: {ods_conv_filepath}");
                        }

                        // Ordinary use, no archiving
                        else
                        {
                            numCOMPLETE++;
                            // Delete copied spreadsheet
                            string new_location = docCollection + "\\" + Path.GetFileName(conv_filepath);
                            File.Move(conv_filepath, new_location);
                            if (File.Exists(copy_filepath))
                            {
                                File.Delete(copy_filepath);
                            }
                            Directory.Delete(file_folder);
                            copy_extension = "";
                            copy_filename = "";
                            copy_filepath = "";
                            xlsx_conv_extension = ".xlsx";
                            xlsx_conv_filename = Path.GetFileName(conv_filepath);
                            xlsx_conv_filepath = conv_filepath;
                            ods_conv_extension = null;
                            ods_conv_filename = null;
                            ods_conv_filepath = null;
                            // Inform user
                            Console.WriteLine($"--> File saved to: {conv_filepath}");
                        }
                    }
                    else
                    {
                        convert_success = false;
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
