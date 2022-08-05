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
        // Define data types
        public static int numCOMPLETE = 0;
        public static int numFAILED = 0;
        public static int numXLSX_noconversion = 0;
        public static int numODS_noconversion = 0;
        public static int numTOTAL_conv = numCOMPLETE + numXLSX_noconversion;
        public static bool? convert_success = null;
        static string? file_folder = null;
        static int subdir_number = 1;
        static int copy_file_number = 1;
        static int conv_file_number = 1;
        static string? copy_extension = null;
        static string? copy_filename = null;
        static string? copy_filepath = null;
        static string? conv_extension = null;
        static string? conv_filename = null;
        static string? conv_filepath = null;
        static string? error_message = null;
        static string[] error_messages = { "", "Legacy Excel file formats are not supported", "Binary .xlsb file format needs Excel installed with .NET programming", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format cannot contain any cell values and is not converted", "Spreadsheet is already .xlsx file format", "Spreadsheet cannot be opened, because the XML structure is malformed", "Spreadsheet was converted to OOXML Transitional conformance", ".xlsx Strict conformance identified", "Cannot convert automatically because of irregular content", "Google Sheets are stored in cloud and cannot be converted locally", "Apple Numbers are not supported", "Converted to Strict conformance", "Microsoft Excel is not installed. Excel is necessary to convert conformance from Transitional to Strict" };

        // Convert spreadsheets method
        public List<fileIndex> Convert_Spreadsheets(string function, string inputdir, bool recurse, string Results_Directory)
        {
            Console.WriteLine("---");
            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original file format;Convert filepath;Convert filename;Convert file format;Convert success;Convert Message");
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
                string no_ext = Path.GetFileNameWithoutExtension(copy_filepath);

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

                        // OpenDocument file formats using LibreOffice
                        case ".fods":
                        case ".ods":
                        case ".ots":
                        // And binary OOXML
                        case ".xlsb":
                            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified
                            // Conversion code
                            convert_success = Convert_from_OpenDocument(function, copy_filepath, file_folder);
                            if (convert_success == true)
                            {
                                conv_extension = ".xlsx";
                                conv_filename = Path.GetFileNameWithoutExtension(copy_filename) + conv_extension;
                                conv_filepath = docCollection + "\\" + conv_filename;
                                // Prevent overriding of existing conversion when moving to docCollection
                                while (File.Exists(conv_filepath))
                                {
                                    conv_file_number++;
                                    conv_filepath = docCollection + "\\" + no_ext + "_" + conv_file_number + conv_extension;
                                }
                                File.Move(copy_filepath, conv_filepath);
                                File.Delete(file_folder + "\\" + no_ext + ".xlsx");
                                numCOMPLETE++;
                            }
                            // If OpenDocument spreadsheet is password protected or corrupt
                            else if (convert_success == false)
                            {

                                Console.WriteLine($"--> Inform something is wrong");
                            }
                            break;

                        // Microsoft Excel Add-in file formats are not converted
                        case ".xla":
                        case ".xlam":
                            // Transform data types
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[5];
                            conv_extension = null;
                            conv_filename = null;
                            conv_filepath = null;
                            break;

                        // Legacy Microsoft Excel file formats
                        case ".xls":
                        case ".xlt":
                            // Transform data types for converted spreadsheets
                            conv_extension = ".xlsx";
                            conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                            conv_filepath = docCollection + "\\" + conv_filename;
                            // Prevent overriding of existing conversion when converting
                            while (File.Exists(conv_filepath))
                            {
                                conv_file_number++;
                                conv_filepath = docCollection + "\\" + no_ext + "_" + conv_file_number + conv_extension;
                            }
                            // Conversion code
                            convert_success = Convert_from_LegacyExcel(org_filepath, copy_filepath, conv_filepath);
                            numCOMPLETE++;
                            break;

                        case ".xlsm":
                        case ".xltm":
                        case ".xltx":
                            // Transform data types for converted spreadsheets
                            conv_extension = ".xlsx";
                            conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                            conv_filepath = docCollection + "\\" + conv_filename;
                            // Prevent overriding of existing conversion when converting
                            while (File.Exists(conv_filepath))
                            {
                                conv_file_number++;
                                conv_filepath = docCollection + "\\" + no_ext + "_" + conv_file_number + conv_extension;
                            }
                            // Conversion code
                            convert_success = Convert_to_OOXML_Transitional(copy_filepath, conv_filepath);
                            numCOMPLETE++;
                            break;

                        case ".xlsx":
                            try
                            {
                                // Open to find Strict conformance
                                SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(org_filepath, false);
                                bool? strict = spreadsheet.StrictRelationshipFound;
                                spreadsheet.Close();
                                if (strict == true)
                                {
                                    // Create data types for converted spreadsheets
                                    conv_extension = ".xlsx";
                                    conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                                    conv_filepath = docCollection + "\\" + conv_filename;
                                    // Prevent overriding of existing conversion when converting
                                    while (File.Exists(conv_filepath))
                                    {
                                        conv_file_number++;
                                        conv_filepath = docCollection + "\\" + no_ext + "_" + conv_file_number + conv_extension;
                                    }
                                    error_message = error_messages[8];
                                    // Conversion code
                                    convert_success = Convert_Strict_to_Transitional(copy_filepath, conv_filepath, file_folder);
                                    numCOMPLETE++;
                                }
                                else
                                {
                                    // No conversion
                                    // Transform data types
                                    numXLSX_noconversion++;
                                    convert_success = false;
                                    error_message = error_messages[6];
                                    conv_extension = null;
                                    conv_filename = null;
                                    conv_filepath = null;
                                }
                            }
                            // If spreadsheet is password protected or corrupt
                            catch (FileFormatException)
                            {
                                // Transform data types
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[4];
                                conv_extension = null;
                                conv_filename = null;
                                conv_filepath = null;
                            }
                            catch (OpenXmlPackageException)
                            {
                                // Transform data types
                                numFAILED++;
                                convert_success = false;
                                conv_extension = null;
                                conv_filename = null;
                                conv_filepath = null;
                                error_message = error_messages[4];
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
                    conv_extension = null;
                    conv_filename = null;
                    conv_filepath = null;
                }
                catch (InvalidDataException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                    conv_extension = null;
                    conv_filename = null;
                    conv_filepath = null;
                }
                // If file is corrupt and cannot be opened for XML schema validation
                catch (OpenXmlPackageException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[7];
                    conv_extension = null;
                    conv_filename = null;
                    conv_filepath = null;
                }
                // If LibreOffice is not installed
                catch (Win32Exception)
                {
                    // COde to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[3];
                    conv_extension = null;
                    conv_filename = null;
                    conv_filepath = null;
                }
                // NPOI encryption
                catch (NPOI.Util.RecordFormatException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[4];
                    conv_extension = null;
                    conv_filename = null;
                    conv_filepath = null;
                }
                // NPOI can't handle old Excel formats in BIFF format
                catch (NPOI.HSSF.OldExcelFormatException)
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
                // NPOI creates this system exception
                catch (NotImplementedException)
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
                // NPOI exception because of formula range with unused values
                catch (NPOI.SS.Formula.FormulaParseException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[10];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }
                // Another NPOI. Try using libreOffice in the catch
                catch (System.InvalidOperationException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[10];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }
                // Another NPOI but this one gives a generic system exception - Dangerous to catch it here, because it could be used in other contexts
                catch (System.IndexOutOfRangeException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;
                    error_message = error_messages[10];
                    xlsx_conv_extension = null;
                    xlsx_conv_filename = null;
                    xlsx_conv_filepath = null;
                    ods_conv_extension = null;
                    ods_conv_filename = null;
                    ods_conv_filepath = null;
                }

                finally
                {
                    // Delete copied spreadsheet
                    File.Delete(copy_filepath);
                    Directory.Delete(file_folder);
                    // Delete info of copied spreadsheet
                    copy_extension = null;
                    copy_filename = null;
                    copy_filepath = null;

                    // Inform user
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Conversion {convert_success}");
                    if (convert_success == true)
                    {
                        Console.WriteLine($"--> Conversion saved to: {conv_filepath}");
                    }
                    else if (error_message != null || error_message == error_messages[6])
                    {
                        Console.WriteLine(error_message);
                    }
                    Console.WriteLine("---");

                    // Add copied and converted spreadsheets file info to index of files
                    File_List.Add(new fileIndex { File_Folder = file_folder, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = copy_filepath, Copy_Filename = copy_filename, Copy_Extension = copy_extension, Conv_Filepath = conv_filepath, Conv_Filename = conv_filename, Conv_Extension = conv_extension, Convert_Success = convert_success });

                    // Output result in open CSV file
                    var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{conv_filepath};{conv_filename};{conv_extension};{convert_success};{error_message}");
                    csv.AppendLine(newLine2);
                }
            }
            // Close CSV file to log results
            Spreadsheet.CSV_filepath = Results_Directory + "\\2_Convert_Results.csv";
            File.WriteAllText(Spreadsheet.CSV_filepath, csv.ToString());

            // Calculate the number of completed conversions
            numTOTAL_conv = numCOMPLETE + numXLSX_noconversion + numODS_noconversion;

            return File_List;
        }
    }
}
