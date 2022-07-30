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
    public partial class Spreadsheet
    {
        // Public conversion data types
        public static int numCOMPLETE = 0;
        public static int numFAILED = 0;
        public static int numXLSX_noconversion = 0;
        public static int numTOTAL_conv = numCOMPLETE + numXLSX_noconversion;

        // Convert spreadsheets method
        public List<fileIndex> Convert(string function, string inputdir, bool recurse, string Results_Directory)
        {
            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Create data types
            string? file_folder = null;
            int subdir_number = 1;
            int copy_file_number = 1;
            int conv_file_number = 1;
            string? copy_extension = null;
            string? copy_filename = null;
            string? copy_filepath = null;
            string? conv_extension = null;
            string? conv_filename = null;
            string? conv_filepath = null;
            bool? convert_success = null;
            string docCollection_subdir = "";
            string? error_message = null;
            string[] error_messages = { "", "Legacy Excel file formats are not supported", "Binary .xlsb file format needs Excel installed with .NET programming", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format cannot contain any cell values and is not converted", "Spreadsheet is already .xlsx file format", "Spreadsheet cannot be opened, because the XML structure is malformed", "Spreadsheet was converted to OOXML Transitional conformance" };

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
                docCollection_subdir = docCollection + "\\" + subdir_number;
                while (Directory.Exists(docCollection_subdir))
                {
                     subdir_number++;
                     docCollection_subdir = docCollection + "\\" + subdir_number;
                }
                DirectoryInfo Output_Subdir = Directory.CreateDirectory(docCollection_subdir);

                // Transform data types for copied original spreadsheet
                copy_extension = org_extension;
                copy_filename = "orgFile_" + org_filename;
                copy_filepath = docCollection_subdir + "\\" + copy_filename;

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
                        // And binary OOXML
                        case ".xlsb":
                            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified
                            // Conversion code
                            convert_success = Convert_OpenDocument(copy_filepath, docCollection_subdir);
                            // If archiving, because of previous bug we must rename converted spreadsheet
                            if (convert_success == true && function == "count&convert&compare&archive")
                            {
                                string[] filename = Directory.GetFiles(docCollection_subdir, "*.xlsx");
                                if (filename.Length > 0)
                                {
                                    // Rename converted spreadsheet
                                    string old_filename = filename[0];
                                    string new_filename = file_folder + "\\1.xlsx";
                                    File.Move(old_filename, new_filename);
                                    // Transform datatypes
                                    conv_extension = ".xlsx";
                                    conv_filename = "1.xlsx";
                                    conv_filepath = file_folder + "\\1.xlsx";
                                    numCOMPLETE++;
                                }
                            }
                            // If ordinary use, no archiving
                            else if (convert_success == true)
                            {
                                conv_extension = ".xlsx";
                                conv_filename = Path.GetFileNameWithoutExtension(copy_filename) + conv_extension;
                                string conv_filename_without_ext = Path.GetFileNameWithoutExtension(copy_filename);
                                conv_filepath = docCollection + "\\" + conv_filename;
                                // Prevent overriding of existing conversion when moving to docCollection
                                while (File.Exists(conv_filepath))
                                {
                                    conv_file_number++;
                                    conv_filepath = docCollection + "\\" + conv_filename_without_ext + "_" + conv_file_number + conv_extension;
                                }
                                File.Move(copy_filepath, conv_filepath);
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
                            if (function == "count&convert&compare&archive")
                            {
                                conv_extension = ".xlsx";
                                conv_filename = "1.xlsx";
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                            }
                            // No archiving
                            else
                            {
                                conv_extension = ".xlsx";
                                conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                                conv_filepath = docCollection + "\\" + conv_filename;
                                // Prevent overriding of existing conversion when converting
                                while (File.Exists(conv_filepath))
                                {
                                    conv_file_number++;
                                    conv_filepath = copy_filepath + "_" + conv_file_number;
                                }
                            }
                            // Conversion code
                            convert_success = Convert_Legacy_Excel_NPOI(org_filepath, copy_filepath, conv_filepath);
                            numCOMPLETE++;
                            break;

                        case ".xlsm":
                        case ".xltm":
                        case ".xltx":
                            // Transform data types for converted spreadsheets
                            if (function == "count&convert&compare&archive")
                            {
                                conv_extension = ".xlsx";
                                conv_filename = "1.xlsx";
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                            }
                            // No archiving
                            else
                            {
                                conv_extension = ".xlsx";
                                conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                                conv_filepath = docCollection + "\\" + conv_filename;
                                // Prevent overriding of existing conversion when converting
                                while (File.Exists(conv_filepath))
                                {
                                    conv_file_number++;
                                    conv_filepath = copy_filepath + "_" + conv_file_number;
                                }
                            }
                            // Conversion code
                            convert_success = Convert_OOXML_Transitional(org_filepath, copy_filepath, conv_filepath);
                            numCOMPLETE++;
                            break;

                        case ".xlsx":
                            try
                            {
                                // Open to find Strict conformance
                                SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(org_filepath, false);
                                bool? strict = spreadsheet.StrictRelationshipFound;
                                spreadsheet.Close();
                                // If archiving has been selected
                                if (function == "count&convert&compare&archive")
                                {
                                    // No conversion
                                    // Transform data types
                                    numXLSX_noconversion++;
                                    convert_success = null;
                                    error_message = error_messages[6];
                                    conv_extension = ".xlsx";
                                    conv_filename = "1.xlsx";
                                    conv_filepath = docCollection_subdir + "\\1.xlsx";
                                }
                                // if ordinary usage, no archiving
                                else if (strict == true)
                                {
                                    // Create data types for converted spreadsheets
                                    conv_extension = ".xlsx";
                                    conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                                    conv_filepath = docCollection + "\\" + conv_filename;
                                    // Prevent overriding of existing conversion when converting
                                    while (File.Exists(conv_filepath))
                                    {
                                        conv_file_number++;
                                        conv_filepath = copy_filepath + "_" + conv_file_number;
                                    }
                                    error_message = error_messages[8];
                                    // Conversion code
                                    convert_success = Convert_Transitional_to_Strict(copy_filepath, conv_filepath, file_folder);
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
                finally
                {
                    // Delete copied spreadsheet if no archiving
                    if (function != "count&convert&compare&archive")
                    {
                        File.Delete(copy_filepath); // BUG: Does not delete the file
                        Directory.Delete(docCollection_subdir, true); // To correct bug, written "true" here
                    }
                    // Delete info of copied spreadsheet, if no archiving
                    if (function != "count&convert&compare&archive")
                    {
                        copy_extension = null;
                        copy_filename = null;
                        copy_filepath = null;
                    }
                    // Inform user
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Conversion {convert_success}");
                    if (convert_success == true) 
                    {
                        Console.WriteLine($"--> Conversion saved to: {file_folder}");
                    }
                    else if (error_message != null || error_message == error_messages[6])
                    {
                        Console.WriteLine(error_message);
                    }

                    // Add copied and converted spreadsheets file info to index of files
                    File_List.Add(new fileIndex { File_Folder = docCollection_subdir, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = copy_filepath, Copy_Filename = copy_filename, Copy_Extension = copy_extension, Conv_Filepath = conv_filepath, Conv_Filename = conv_filename, Conv_Extension = conv_extension, Convert_Success = convert_success });

                    // Output result in open CSV file
                    var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{conv_filepath};{conv_filename};{conv_extension};{convert_success};{error_message}");
                    csv.AppendLine(newLine2);
                }
            }
            // Close CSV file to log results
            CSV_filepath = Results_Directory + "\\2_Convert_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Inform user of results
            Convert_Results();

            return File_List;
        }

        public void Convert_Results()
        {
            numTOTAL_conv = numCOMPLETE + numXLSX_noconversion;

            Console.WriteLine("---");
            Console.WriteLine("CONVERT RESULTS");
            Console.WriteLine("---");
            Console.WriteLine($"{numCOMPLETE} out of {numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"{numXLSX_noconversion} spreadsheets were already .xlsx");
            Console.WriteLine($"{numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
            Console.WriteLine("Conversion ended");
            Console.WriteLine("---");
        }
    }
}
