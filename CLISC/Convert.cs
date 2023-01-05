using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    public partial class Conversion
    {
        // Define data types
        public static int numCOMPLETE = 0;
        public static int numFAILED = 0;
        public static bool convert_success = false;
        static string? file_folder = null;
        static int subdir_number = 1;
        static int copy_number = 0;
        string org_extension = "";
        string org_filename = "";
        string org_filepath = "";
        static string? copy_extension = null;
        static string? copy_filename = null;
        static string? copy_filepath = null;
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
        static string[] error_messages = { "", "Spreadsheet cannot be read", "Microsoft Excel Add-In file format cannot contain any cell values and is not converted", "Google Sheets are stored in the cloud and cannot be converted locally", "Apple Numbers file format is not supported", "Filesize exceeds application limit" };

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

                copy_filename = org_filename;
                copy_extension = org_extension;
                copy_filepath = file_folder + "\\" + copy_filename;

                if (function == "CountConvertCompareArchive")
                {
                    // Change data types for copied spreadsheet
                    copy_filename = "orgFile_" + org_filename;
                    copy_filepath = file_folder + "\\" + copy_filename;

                    // Change conversion filepath
                    conv_filepath = file_folder + "\\1.xlsx";
                }
                else
                {
                    // Change conversion filepath
                    conv_filepath = file_folder + "\\" + Path.GetFileNameWithoutExtension(copy_filename) + ".xlsx";
                }

                // Copy spreadsheet
                File.Copy(org_filepath, copy_filepath);

                // Remove file attributes on copied spreadsheet
                File.SetAttributes(copy_filepath, FileAttributes.Normal);

                // Inform user of original filepath
                Console.WriteLine(org_filepath);

                // Convert spreadsheet
                try
                {
                    // Throw exception if filesize is over limit
                    long length = new System.IO.FileInfo(copy_filepath).Length;
                    length = length/1000000;
                    if (length >= 150) // Set limit, currently 150MB
                    {
                        throw new System.Data.ConstraintException("Filesize exceeded");
                    }

                    // Change conversion method based on file extension
                    switch (org_extension)
                    {
                        // Google Sheets file format
                        case ".gsheet":
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[3];
                            break;

                        // OpenDocument file formats and Apple Numbers
                        case ".fods":
                        case ".ods":
                        case ".ots":
                        case ".numbers":
                            // Convert to XLSX Transitional using LibreOffice
                            convert_success = Convert_from_OpenDocument(function, copy_filepath, file_folder);
                            break;

                        // Microsoft Excel Add-in file formats are not converted
                        case ".xla":
                        case ".xlam":
                            // Transform data types
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[2];
                            break;

                        // Legacy Microsoft Excel file formats and OOXML binary
                        case ".xls":
                        case ".xlt":
                        case ".xlsb":
                            // Convert to .xlsx Transitional using Excel Interop
                            convert_success = Convert_Legacy_ExcelInterop(copy_filepath, conv_filepath);
                            break;

                        // Office Open XML file formats
                        case ".xlsm":
                        case ".xlsx":
                        case ".xltm":
                        case ".xltx":
                            // Convert to .xlsx Transitional using Open XML SDK
                            convert_success = Convert_to_OOXML_Transitional(copy_filepath, conv_filepath);
                            break;
                    }
                }

                // Handle any errors occuring during conversion
                catch (FileFormatException) // If spreadsheet is password protected or corrupt
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[1];
                }
                catch (InvalidDataException)
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[1];
                }
                catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException) // If file is corrupt and cannot be opened for XML schema validation
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[1];
                }
                catch (Win32Exception) // If .LibreOffice is not installed
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[1];
                }
                catch (System.Runtime.InteropServices.COMException) // If files used by Excel Interop are password protected or corrupt
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[1];
                }       
                catch (System.Data.ConstraintException) // If filesize exceeds limit
                {
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[5];
                }

                // Post conversion operations
                finally
                {
                    // Inform user
                    Console.WriteLine($"--> Conversion: {convert_success}");

                    // If conversion success
                    if (convert_success == true)
                    {
                        // Count one complete
                        numCOMPLETE++;

                        //Transform XLSX data types
                        xlsx_conv_extension = ".xlsx";
                        xlsx_conv_filename = Path.GetFileName(conv_filepath);
                        xlsx_conv_filepath = conv_filepath;

                        // Inform user
                        Console.WriteLine($"--> File saved to: {conv_filepath}");
                    }
                    // If conversion failed
                    else
                    {
                        if (error_message == null)
                        {
                            error_message = error_messages[1];
                        }
                        Console.WriteLine($"--> {error_message}");
                    }

                    // If archiving, transform data types
                    if (function == "CountConvertCompareArchive")
                    {
                        ods_conv_extension = ".ods";
                        ods_conv_filename = "1.ods";
                        ods_conv_filepath = file_folder + "\\1.ods";
                        error_message = null;

                    }                   
                    // If no archiving
                    else
                    {
                        // Delete the copied spreadsheet, if conversion failed
                        if (!convert_success)
                        {
                            File.Delete(conv_filepath);
                        }
                        // If success, move the spreadsheet to docCollection
                        else
                        {
                            string new_location = docCollection + "\\" + conv_filename;
                            string copy_filename_without_extension = Path.GetFileNameWithoutExtension(copy_filename);
                            while (File.Exists(new_location))
                            {
                                copy_number++;
                                copy_filename = $"{copy_filename_without_extension}({copy_number}){entry.Org_Extension}";
                                copy_filepath = docCollection + "\\" + copy_filename;
                            }
                            File.Move(conv_filepath, new_location);
                        }

                        // Transform data types
                        copy_extension = null;
                        copy_filename = null;
                        copy_filepath = null;

                        // Delete folder
                        Directory.Delete(file_folder);
                    }

                    // Add copied and converted spreadsheets file info to index of files
                    File_List.Add(new fileIndex { File_Folder = file_folder, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = copy_filepath, Copy_Filename = copy_filename, Copy_Extension = copy_extension, XLSX_Conv_Filepath = xlsx_conv_filepath, XLSX_Conv_Filename = xlsx_conv_filename, XLSX_Conv_Extension = xlsx_conv_extension, ODS_Conv_Filepath = ods_conv_filepath, ODS_Conv_Filename = ods_conv_filename, ODS_Conv_Extension = ods_conv_extension, Convert_Success = convert_success });

                    // Output result in open CSV file
                    var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{xlsx_conv_filepath};{ods_conv_filepath};{convert_success};{error_message}");
                    csv.AppendLine(newLine2);
                }
            }
            // Close CSV file to log results
            Results.CSV_filepath = Results_Directory + "\\2_Convert_Results.csv";
            File.WriteAllText(Results.CSV_filepath, csv.ToString(), Encoding.UTF8);

            return File_List;
        }
    }
}
