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
    public partial class Spreadsheet
    {
        // Convert spreadsheets method
        public List<fileIndex> Convert(string argument0, string argument1, string argument3, string Results_Directory)
        {
            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Local conversion error messages
            int numCOMPLETE = 0;
            int numFAILED = 0;
            bool convert_success = false;
            string error_message = "";
            string[] error_messages = { "", "Legacy Excel file formats are not supported", "Binary XLSB file format is not supported", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format is not supported", "Spreadsheet is already .xlsx file format" };

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original file format;Convert filepath;Convert filename;Convert file format;Convert success;Convert Message");
            csv.AppendLine(newLine0);

            // Create lists
            List<fileIndex> Org_File_List = Org_Files(argument1, argument3);
            List<fileIndex> File_List = new List<fileIndex>();

            // Create subdirectory (docCollection) for converted spreadsheet files
            string docCollection = Results_Directory + "\\docCollection";
            DirectoryInfo Output_Dir = Directory.CreateDirectory(docCollection);

            // Convert spreadsheets according to archiving requirements
            if (argument0 == "Count&Convert&Compare&Archive")
            {
                // Loop spreadsheets based on enumeration
                foreach (fileIndex entry in Org_File_List)
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
                    string copy_filepath = docCollection_subdir + "\\" + copy_filename;

                    // Copy spreadsheet
                    File.Copy(org_filepath, copy_filepath);

                    // Create data types for converted spreadsheets
                    string conv_extension = ".xlsx";
                    string conv_filename = "1.xlsx";
                    string conv_filepath = docCollection_subdir + "\\1.xlsx";

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
                                convert_success = Convert_OpenDocument(argument0, copy_filepath, docCollection_subdir);
                                // The next line must exist otherwise CSV will have wrong "conv_new_filepath"
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                break;

                            // Legacy Microsoft Excel file formats
                            case ".xla":
                                // Conversion code
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xls":
                            case ".xlt":
                                // Conversion code
                                convert_success = Convert_Legacy_Excel_NPOI(org_filepath, conv_filepath);
                                break;

                            // Office Open XML file formats
                            case ".xlsb":
                                // Conversion code
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[2];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xlam":
                                // Conversion code
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xlsm":
                            case ".xltm":
                            case ".xltx":
                                // The next line must exist otherwise the switch will not convert
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                // Conversion code
                                convert_success = Convert_OOXML_Transitional(copy_filepath, conv_filepath);
                                break;

                            case ".xlsx":
                                // No converison
                                convert_success = true;
                                error_message = error_messages[6];

                                // Rename copied spreadsheet
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                File.Move(copy_filepath, conv_filepath);

                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
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
                        conv_extension = "";
                        conv_filename = "";
                        conv_filepath = "";
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }

                    // If LibreOffice is not installed
                    catch (Win32Exception)
                    {
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[3];
                        conv_extension = "";
                        conv_filename = "";
                        conv_filepath = "";
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }

                    // NPOI encryption
                    catch(NPOI.Util.RecordFormatException)
                    {
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[4];
                        conv_extension = "";
                        conv_filename = "";
                        conv_filepath = "";
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
                foreach (fileIndex entry in Org_File_List)
                {
                    // Find file information
                    string org_extension = entry.Org_Extension;
                    string org_filename = entry.Org_Filename;
                    string org_filepath = entry.Org_Filepath;
                    string conv_extension = ".xlsx";
                    string conv_filename = Path.GetFileNameWithoutExtension(entry.Org_Filename) + conv_extension;
                    string conv_filepath = docCollection + "\\" + conv_filename;

                    // Try to convert spreadsheet
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
                                convert_success = Convert_OpenDocument(org_filepath, conv_filepath, docCollection);
                                break;

                            // Legacy Microsoft Excel file formats
                            case ".xla":
                                // Conversion code
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xls":
                            case ".xlt":
                                // Conversion code
                                //convert_success = Convert_Legacy_Excel_NPOI(org_filepath);
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[1];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            // Office Open XML file formats
                            case ".xlsb":
                                // Conversion code
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[2];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xlam":
                                // Conversion code
                                numFAILED++;
                                convert_success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                                break;

                            case ".xlsm":
                            case ".xltm":
                            case ".xltx":
                                // Conversion code
                                convert_success = Convert_OOXML_Transitional(org_filepath, conv_filepath);
                                break;

                            case ".xlsx":
                                // No converison
                                convert_success = false;
                                error_message = error_messages[6];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(org_filepath);
                                Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
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
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }
                    // If LibreOffice is not installed
                    catch (Win32Exception)
                    {
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[3];
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }

                    // NPOI encryption
                    catch (NPOI.Util.RecordFormatException)
                    {
                        numFAILED++;
                        convert_success = false;
                        error_message = error_messages[4];
                        conv_extension = "";
                        conv_filename = "";
                        conv_filepath = "";
                        // Inform user
                        Console.WriteLine(org_filepath);
                        Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                    }

                    finally
                    {
                        File_List.Add(new fileIndex{File_Folder = docCollection, Org_Filepath = org_filepath, Org_Filename = org_filename, Org_Extension = org_extension, Copy_Filepath = "", Copy_Filename = "", Copy_Extension = "", Conv_Filepath = conv_filepath, Conv_Filename = conv_filename, Conv_Extension = conv_extension, Convert_Success = convert_success});
                        // Output result in open CSV file
                        var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{conv_filepath};{conv_filename};{conv_extension};{convert_success};{error_message}");
                        csv.AppendLine(newLine2);
                    }
                }
            }

            // Close CSV file to log results
            string CSV_filepath = Results_Directory + "\\2_Convert_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Inform user of results
            numCOMPLETE = numTOTAL - numFAILED;
            Console.WriteLine("---");
            Console.WriteLine($"{numCOMPLETE} out of {numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"{numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
            Console.WriteLine("Conversion finished");
            Console.WriteLine("---");

            return File_List;
        }
    }
}
