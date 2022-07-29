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

            // Local conversion error messages

            bool? convert_success = null;
            string error_message = "";
            string[] error_messages = { "", "Legacy Excel file formats are not supported", "Binary XLSB file format is not supported", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format cannot contain any cell values and is not converted", "Spreadsheet is already .xlsx file format", "Spreadsheet cannot be opened, because the XML structure is malformed" };

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

            string? copy_filepath = null;
            string? conv_extension = null;
            string? conv_filename = null;
            string? conv_filepath = null;

            // Convert spreadsheets according to archiving requirements
            if (function == "count&convert&compare&archive")
            {
                // Loop spreadsheets based on enumeration
                foreach (var entry in Org_File_List)
                {
                    // Create new subdirectory for the spreadsheet
                    int subdir_number = 1;
                    string docCollection_subdir = docCollection + "\\" + subdir_number;
                    while (Directory.Exists(docCollection_subdir))
                    {
                        subdir_number++;
                        docCollection_subdir = docCollection + "\\" + subdir_number;
                    }
                    DirectoryInfo Output_Subdir = Directory.CreateDirectory(docCollection_subdir);

                    // Find file information
                    string org_extension = entry.Org_Extension;
                    string org_filename = entry.Org_Filename;
                    string org_filepath = entry.Org_Filepath;

                    // Create data types for copied original spreadsheets
                    string copy_extension = org_extension;
                    string copy_filename = "orgFile_" + org_filename;
                    copy_filepath = docCollection_subdir + "\\" + copy_filename;

                    // Copy spreadsheet
                    File.Copy(org_filepath, copy_filepath);

                    // Convert spreadsheet
                    try
                    {
                        // Change conversion method based on file extension
                        switch (org_extension)
                        {
                            // OpenDocument file formats
                            case ".fods":
                            case ".ods":
                            case ".ots":
                                // Conversion code
                                convert_success = Convert_OpenDocument(function, org_filepath, copy_filepath, docCollection_subdir);
                                error_message = "";
                                // Transform data types for converted spreadsheets - code must be here because of bug, where a password protected ods file won't be caught
                                if (convert_success == true)
                                {
                                    conv_extension = ".xlsx";
                                    conv_filename = "1.xlsx";
                                    conv_filepath = docCollection_subdir + "\\1.xlsx";
                                    error_message = "";
                                    numCOMPLETE++;
                                }
                                else
                                {
                                    error_message = "Spreadsheet is password protected or corrupt";
                                }
                                break;

                            // Legacy Microsoft Excel file formats
                            case ".xla":
                                // Transform data types
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = null;
                                conv_filename = null;
                                conv_filepath = null;
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xls":
                            case ".xlt":
                                // Transform data types for converted spreadsheets
                                conv_extension = ".xlsx";
                                conv_filename = "1.xlsx";
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                // Conversion code
                                convert_success = Convert_Legacy_Excel_NPOI(org_filepath, copy_filepath, conv_filepath);
                                numCOMPLETE++;
                                error_message = "";
                                break;

                            // Office Open XML file formats
                            case ".xlsb":
                                // Conversion code using Excel
                                convert_success = Convert_XLSB(org_filepath, copy_filepath, conv_filepath);
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success}");
                                numCOMPLETE++;
                                error_message = "";
                                break;

                            case ".xlam":
                                // Transform data types
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = null;
                                conv_filename = null;
                                conv_filepath = null;
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xlsm":
                            case ".xltm":
                            case ".xltx":
                                // Transform data types for converted spreadsheets
                                conv_extension = ".xlsx";
                                conv_filename = "1.xlsx";
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                // Conversion code
                                convert_success = Convert_OOXML_Transitional(org_filepath, copy_filepath, conv_filepath);
                                numCOMPLETE++;
                                error_message = "";
                                break;

                            case ".xlsx":
                                try
                                {
                                    // No converison
                                    // Try to open spreadsheet to cause catch if password protected or corrupt
                                    SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(copy_filepath, false);
                                    spreadsheet.Close();
                                    // Transform data types for converted spreadsheets
                                    numXLSX_noconversion++;
                                    conv_extension = ".xlsx";
                                    conv_filename = "1.xlsx";
                                    conv_filepath = docCollection_subdir + "\\1.xlsx";
                                    convert_success = null;
                                    error_message = error_messages[6];
                                    // Copy spreadsheet again and rename copy
                                    File.Copy(org_filepath, conv_filepath);
                                    // Inform user
                                    Console.WriteLine(org_filepath);
                                    Console.WriteLine($"--> {error_message}");
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
                                    // Inform user
                                    Console.WriteLine(org_filepath);
                                    Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
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
                                    // Inform user
                                    Console.WriteLine(org_filepath);
                                    Console.WriteLine($"--> {error_message}");
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
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
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
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
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
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
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
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }
                    // NPOI encryption
                    catch(NPOI.Util.RecordFormatException)
                    {
                        // Code to execute
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[4];
                        conv_extension = null;
                        conv_filename = null;
                        conv_filepath = null;
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }

                    finally
                    {
                        // Add copied and converted spreadsheets file info to index of files
                        File_List.Add(new fileIndex{File_Folder = docCollection_subdir, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = copy_filepath, Copy_Filename = copy_filename, Copy_Extension = copy_extension, Conv_Filepath = conv_filepath, Conv_Filename = conv_filename, Conv_Extension = conv_extension, Convert_Success = convert_success});

                        // Output result in open CSV file
                        var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{conv_filepath};{conv_filename};{conv_extension};{convert_success};{error_message}");
                        csv.AppendLine(newLine2);
                    }
                }
            }

            // Convert spreadsheets ordinarily, no archiving
            else
            {
                // Loop spreadsheets based on enumeration
                foreach (var entry in Org_File_List)
                {
                    // Find file information
                    string org_extension = entry.Org_Extension;
                    string org_filename = entry.Org_Filename;
                    string org_filepath = entry.Org_Filepath;

                    // Try to convert spreadsheet
                    try
                    {
                        // Create data types for converted spreadsheets
                        conv_extension = ".xlsx";
                        conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                        conv_filepath = docCollection + "\\" + conv_filename;

                        // Change conversion method based on file extension
                        switch (org_extension)
                        {
                            // OpenDocument file formats
                            case ".fods":
                            case ".ods":
                            case ".ots":
                                // Conversion code
                                convert_success = Convert_OpenDocument(function, org_filepath, copy_filepath, docCollection);
                                numCOMPLETE++;
                                error_message = "";
                                break;

                            // Legacy Microsoft Excel file formats
                            case ".xla":
                                // Transform data types
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = null;
                                conv_filename = null;
                                conv_filepath = null;
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xls":
                            case ".xlt":
                                // Conversion code
                                convert_success = Convert_Legacy_Excel_NPOI(copy_filepath, org_filepath, conv_filepath);
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success}");
                                numCOMPLETE++;
                                error_message = "";
                                break;

                            // Office Open XML file formats
                            case ".xlsb":
                                // Conversion code using Excel
                                convert_success = Convert_XLSB(org_filepath, copy_filepath, conv_filepath);
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success}");
                                numCOMPLETE++;
                                error_message = "";
                                break;

                            case ".xlam":
                                // Transform data types
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = null;
                                conv_filename = null;
                                conv_filepath = null;
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xlsm":
                            case ".xltm":
                            case ".xltx":
                                // Conversion code
                                convert_success = Convert_OOXML_Transitional(copy_filepath, org_filepath, conv_filepath);
                                numCOMPLETE++;
                                error_message = "";
                                break;

                            case ".xlsx":
                                // Conversion code for Transitional to Strict
                                SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(org_filepath, false);
                                bool? strict = spreadsheet.StrictRelationshipFound;
                                spreadsheet.Close();
                                if (strict == true)
                                {
                                    // Conversion code
                                    convert_success = Convert_OOXML_Transitional(copy_filepath, org_filepath, conv_filepath);
                                    numCOMPLETE++;
                                    error_message = "";
                                    // Inform user
                                    Console.WriteLine(org_filepath);
                                    Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                }
                                else
                                {
                                    // Transform data types
                                    numXLSX_noconversion++;
                                    convert_success = false;
                                    error_message = error_messages[6];
                                    conv_extension = null;
                                    conv_filename = null;
                                    conv_filepath = null;
                                    // Inform user
                                    Console.WriteLine(org_filepath);
                                    Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                }
                                break;
                        }
                    }

                    // If spreadsheet is password protected or corrupt
                    catch (FileFormatException)
                    {
                        // Transform data types
                        numFAILED++;
                        convert_success = false;
                        conv_filepath = null;
                        conv_filename = null;
                        conv_extension = null;
                        error_message = error_messages[4];
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }
                    catch (InvalidDataException)
                    {
                        // Transform data types
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[4];
                        conv_extension = null;
                        conv_filename = null;
                        conv_filepath = null;
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }
                    // If file is corrupt and cannot be opened for XML schema validation
                    catch (OpenXmlPackageException)
                    {
                        // Transform data types
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[7];
                        conv_extension = null;
                        conv_filename = null;
                        conv_filepath = null;
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }
                    // If LibreOffice is not installed
                    catch (Win32Exception)
                    {
                        // Transform data types
                        numFAILED++;
                        convert_success = false;
                        conv_filepath = null;
                        conv_filename = null;
                        conv_extension = null;
                        error_message = error_messages[3];
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }

                    // NPOI has special exception for handling password protected or corrupt files
                    catch (NPOI.Util.RecordFormatException)
                    {
                        // Transform data types
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[4];
                        conv_filepath = null;
                        conv_filename = null;
                        conv_extension = null;
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }

                    finally
                    {
                        // Add file to fileIndex of File_List
                        File_List.Add(new fileIndex{File_Folder = docCollection, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = null, Copy_Filename = null, Copy_Extension = null, Conv_Filepath = conv_filepath, Conv_Filename = conv_filename, Conv_Extension = conv_extension, Convert_Success = convert_success});

                        // Output result in open CSV file
                        var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{conv_filepath};{conv_filename};{conv_extension};{convert_success};{error_message}");
                        csv.AppendLine(newLine2);
                    }
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
