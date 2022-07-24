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
        // Create data types
        public bool Convert_Success;

        // Convert spreadsheets method
        public List<fileIndex> Convert(string argument0, string argument1, string argument3, string Results_Directory)
        {
            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Local conversion error messages
            int numCOMPLETE = 0;
            int numFAILED = 0;
            string error_message = "";
            string[] error_messages = { "", "Legacy Excel file formats are not supported", "Binary XLSB file format is not supported", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format is not supported", "Spreadsheet is already .xlsx file format" };

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original file format;Convert filepath;Convert filename;Convert file format;Convert success;Convert Message");
            csv.AppendLine(newLine0);

            // Create enumeration of original spreadsheets based on input directory
            List<fileIndex> File_List = fileIndex.Org_Files(argument1, argument3);

            // Create subdirectory (docCollection) for converted spreadsheet files
            string docCollection = Results_Directory + "\\docCollection";
            DirectoryInfo Output_Dir = Directory.CreateDirectory(docCollection);

            // Convert spreadsheets according to archiving requirements
            if (argument0 == "Count&Convert&Compare&Archive")
            {
                // Loop spreadsheets based on enumeration
                foreach (var file in File_List)
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

                    // Create data types for copied original spreadsheets
                    string copy_extension = this._Org_Extension;
                    string copy_filename = "orgFile_" + this._Org_Filename;
                    string copy_filepath = docCollection_subdir + "\\" + copy_filename;

                    // Copy spreadsheet
                    File.Copy(this._Org_Filepath, copy_filepath);

                    // Create data types for converted spreadsheets
                    string conv_extension = ".xlsx";
                    string conv_filename = "1" + conv_extension;
                    string conv_filepath = docCollection_subdir + "\\" + conv_filename;

                    // Convert spreadsheet
                    try
                    {
                        // Change conversion method based on file extension
                        switch (this._Org_Extension)
                        {
                            // OpenDocument file formats
                            case ".fods":
                            case ".ods":
                            case ".ots":
                                // Conversion code
                                Convert_Success = Convert_OpenDocument(argument0, copy_filepath, docCollection_subdir);
                                // The next line must exist otherwise CSV will have wrong "conv_new_filepath"
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                break;

                            // Legacy Microsoft Excel file formats
                            case ".xla":
                                // Conversion code
                                numFAILED++;
                                Convert_Success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;

                            case ".xls":
                            case ".xlt":
                                // Conversion code
                                Convert_Success = Convert_Legacy_Excel_NPOI(this._Org_Filepath, conv_filepath);
                                break;

                            // Office Open XML file formats
                            case ".xlsb":
                                // Conversion code
                                numFAILED++;
                                Convert_Success = false;
                                error_message = error_messages[2];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;

                            case ".xlam":
                                // Conversion code
                                numFAILED++;
                                Convert_Success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;

                            case ".xlsm":
                            case ".xltm":
                            case ".xltx":
                                // The next line must exist otherwise the switch will not convert
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                // Conversion code
                                Convert_Success = Convert_OOXML_Transitional(copy_filepath, conv_filepath);
                                break;

                            case ".xlsx":
                                // No converison
                                Convert_Success = true;
                                error_message = error_messages[6];

                                // Rename copied spreadsheet
                                conv_filepath = docCollection_subdir + "\\1.xlsx";
                                File.Move(copy_filepath, conv_filepath);

                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;
                        }
                    }

                    // If spreadsheet is password protected or corrupt
                    catch (FileFormatException)
                    {
                        // Code to execute
                        numFAILED++;
                        Convert_Success = false;
                        error_message = error_messages[4];
                        conv_extension = "";
                        conv_filename = "";
                        conv_filepath = "";
                        // Inform user
                        Console.WriteLine(this._Org_Filepath);
                        Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                    }

                    // If LibreOffice is not installed
                    catch (Win32Exception)
                    {
                        numFAILED++;
                        Convert_Success = false;
                        error_message = error_messages[3];
                        conv_extension = "";
                        conv_filename = "";
                        conv_filepath = "";
                        // Inform user
                        Console.WriteLine(this._Org_Filepath);
                        Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                    }

                    finally
                    {
                        // Add copied and converted spreadsheets file info to index of files
                        File_List.Add(new Spreadsheet(docCollection_subdir, this.Org_Filepath, this.Org_Filename, this._Org_Extension, copy_filepath, copy_filename, copy_extension, conv_filepath, conv_filename, conv_extension, Convert_Success));
                        // Output result in open CSV file
                        var newLine2 = string.Format($"{this.Org_Filepath};{this.Org_Filename};{this._Org_Extension};{conv_filepath};{conv_filename};{conv_extension};{Convert_Success};{error_message}");
                        csv.AppendLine(newLine2);
                    }
                }
            }

            // Convert spreadsheets ordinarily, no archiving
            else
            {
                // Loop spreadsheets based on enumeration
                foreach (var file in File_List)
                {
                    // Create instance for finding file information
                    FileInfo file_info = new FileInfo(file);

                    // Create data types for fileinfo
                    string conv_extension = ".xlsx";
                    string conv_filename = Path.GetFileNameWithoutExtension(file_info.Name) + conv_extension;
                    string conv_filepath = docCollection + "\\" + conv_filename;

                    // Try to convert spreadsheet
                    try
                    {
                        // Change conversion method based on file extension
                        switch (this._Org_Extension)
                        {
                            // OpenDocument file formats
                            case ".fods":
                            case ".ods":
                            case ".ots":
                                // Conversion code
                                Convert_Success = Convert_OpenDocument(this._Org_Filepath, conv_filepath, docCollection);
                                break;

                            // Legacy Microsoft Excel file formats
                            case ".xla":
                                // Conversion code
                                numFAILED++;
                                Convert_Success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;

                            case ".xls":
                            case ".xlt":
                                // Conversion code
                                //convert_success = Convert_Legacy_Excel_NPOI(org_filepath);
                                numFAILED++;
                                Convert_Success = false;
                                error_message = error_messages[1];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;

                            // Office Open XML file formats
                            case ".xlsb":
                                // Conversion code
                                numFAILED++;
                                Convert_Success = false;
                                error_message = error_messages[2];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;

                            case ".xlam":
                                // Conversion code
                                numFAILED++;
                                Convert_Success = false;
                                error_message = error_messages[5];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;

                            case ".xlsm":
                            case ".xltm":
                            case ".xltx":
                                // Conversion code
                                Convert_Success = Convert_OOXML_Transitional(this._Org_Filepath, conv_filepath);
                                break;

                            case ".xlsx":
                                // No converison
                                Convert_Success = false;
                                error_message = error_messages[6];
                                conv_extension = "";
                                conv_filename = "";
                                conv_filepath = "";
                                // Inform user
                                Console.WriteLine(this._Org_Filepath);
                                Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                                break;
                        }
                    }

                    // If spreadsheet is password protected or corrupt
                    catch (FileFormatException)
                    {
                        // Code to execute
                        numFAILED++;
                        Convert_Success = false;
                        error_message = error_messages[4];
                        // Inform user
                        Console.WriteLine(this._Org_Filepath);
                        Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                    }
                    // If LibreOffice is not installed
                    catch (Win32Exception)
                    {
                        numFAILED++;
                        Convert_Success = false;
                        error_message = error_messages[3];
                        // Inform user
                        Console.WriteLine(this._Org_Filepath);
                        Console.WriteLine($"--> Conversion {Convert_Success} - {error_message}");
                    }

                    finally
                    {
                        // Add converted spreadsheets file info to index of files
                        File_List.Add(new Spreadsheet("", this.Org_Filepath, this.Org_Filename, this._Org_Extension, "", "", "", conv_filepath, conv_filename, conv_extension, Convert_Success));
                        // Output result in open CSV file
                        var newLine2 = string.Format($"{this._Org_Filepath};{this._Org_Filename};{this._Org_Extension};{conv_filepath};{conv_filename};{conv_extension};{Convert_Success};{error_message}");
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
