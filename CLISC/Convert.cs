using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop; // not used
using Excel = Microsoft.Office.Interop.Excel; // not used
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{

    public partial class Spreadsheet
    {
        // Create public conversion error messages
        public bool convert_success;
        public string error_message = "";

        // Convert spreadsheets method
        public void Convert(string argument1, string results_directory, string argument3, string argument4)
        {

            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Local conversion error messages
            int numCOMPLETE = 0;
            int numFAILED = 0;

            string[] error_messages = { "", "Legacy Excel file formats are not supported", "Binary XLSB file format is not supported", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format is not supported", "Spreadsheet is already .xlsx file format. File was copied and renamed" };

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original file format;Convert filepath;Convert filename;Convert file format;Convert success;Convert Message");
            csv.AppendLine(newLine0);

            // Create subdirectory for spreadsheet files
            string file_dir = results_directory + "\\docCollection\\";
            DirectoryInfo Output_Dir = Directory.CreateDirectory(@file_dir);

            // Create data types for converted spreadsheets
            int subdir_number = 1;
            string file_subdir = file_dir + subdir_number;
            int conv_file_number = 1;
            string conv_extension = ".xlsx";
            string conv_filename = conv_file_number + conv_extension;
            string conv_filepath = file_subdir + "\\" + conv_filename;

            // Create enumeration of original spreadsheets based on input directory
            List<string> org_enumeration = Enumerate_Original(argument1, argument3);

            // Loop spreadsheets based on enumeration
            foreach (var file in org_enumeration.ToList()) // Is .ToList() necessary?
            {

                // Create instance for finding file information
                FileInfo file_info = new FileInfo(file);

                // Combine data types to original spreadsheets
                org_extension = file_info.Extension;
                org_filename = file_info.Name;
                org_filepath = file_info.FullName;

                // Create data types for copied original spreadsheets
                //string copy_extension = "orgFile_" + org_filename;
                string copy_filename = "orgFile_" + org_filename;
                string copy_filepath = file_subdir + "\\" + copy_filename;

                // Create new subdirectory for the spreadsheet
                while (Directory.Exists(@file_subdir))
                {
                    subdir_number++;
                    file_subdir = file_dir + subdir_number;
                }
                DirectoryInfo Output_Subdir = Directory.CreateDirectory(@file_subdir);

                // Copy spreadsheet
                copy_filepath = file_subdir + "\\" + "orgFile_" + org_filename;
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
                            convert_success = Convert_OpenDocument(org_filepath, copy_filepath, file_subdir);

                            // The next line must exist otherwise CSV will have wrong "conv_new_filepath"
                            conv_filepath = file_subdir + "\\" + conv_file_number + ".xlsx";
                            break;

                        // Legacy Microsoft Excel file formats
                        case ".xla":
                            // Conversion code
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[5];

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                            break;

                        case ".xls":
                        case ".xlt":
                            // Conversion code
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[1];

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

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                            break;

                        case ".xlam":
                            // Conversion code
                            numFAILED++;
                            convert_success = false;
                            error_message = error_messages[5];

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                            break;

                        case ".xlsm":
                        case ".xltm":
                        case ".xltx":
                            // The next line must exist otherwise the switch will not convert
                            conv_filepath = file_subdir + "\\" + conv_file_number + ".xlsx";

                            // Conversion code
                            convert_success = Convert_OOXML(copy_filepath);

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success}");
                            break;

                        case ".xlsx":
                            // No converison
                            convert_success = true;
                            error_message = error_messages[6];

                            // Copy spreadsheet
                            conv_filepath = file_subdir + "\\" + conv_file_number + ".xlsx";
                            File.Copy(copy_filepath, conv_filepath);

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

                    // Output result in open CSV file
                    var newLine1 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_filepath};{copy_filename};;;;{convert_success};{error_message}");
                    csv.AppendLine(newLine1);
                }

                // If LibreOffice is not installed
                catch (Win32Exception)
                {
                    numFAILED++;
                    convert_success = false;
                    string error_message = error_messages[3];

                    // Inform user
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Conversion {convert_success} - {error_message}");
                }

                finally
                {
                    // Output result in open CSV file
                    var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_filepath};{copy_filename};;;;{convert_success};{error_message}");
                    csv.AppendLine(newLine2);
                }

            }

            // Close CSV file to log results
            string CSV_filepath = results_directory + "\\2_Convert_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());
            
            // Inform user of results
            numCOMPLETE = numTOTAL - numFAILED;
            Console.WriteLine("---");
            Console.WriteLine($"{numCOMPLETE} out of {numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"{numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}");
            Console.WriteLine("Conversion finished");
            Console.WriteLine("---");

        }

    }

}
