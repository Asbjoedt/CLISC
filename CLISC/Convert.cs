using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop; // not used
using Excel = Microsoft.Office.Interop.Excel; // not used

namespace CLISC
{

    public partial class Spreadsheet
    {

        // Convert spreadsheets
        public void Convert(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Conversion error messages
            bool success;
            string[] error_message = { "", "Legacy Excel file formats are not supported", "Binary XLSB file format is not supported", "OpenDocument file formats are not supported", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format is not supported", "Spreadsheet is already .xlsx file format. File was copied and renamed" };

            // Open CSV file to log results
            int numFAILED = 0;
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath,Original filename,Original file format,New copy filepath,New copy filename, New convert filepath, New convert filename, New convert file format,Success,Message");
            csv.AppendLine(newLine0);

            // Create subdirectory
            int results_directory_number = 1;
            string results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
            while (Directory.Exists(@results_directory))
            {
                results_directory_number++;
                results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
            }
            results_directory_number = results_directory_number - 1;
            results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
            string conv_dir = results_directory + "\\docCollection\\";
            DirectoryInfo OutputDir = Directory.CreateDirectory(@conv_dir);

            // Copy spreadsheets to subdirectory recursively
            if (argument3 == "Recursive=Yes")
            {
                var extensions = new List<string> { ".fods", ".ods", ".ots", ".xla", ".xls", ".xlt", ".xlam", ".xlsb", ".xlsm", ".xlsx", ".xltm", ".xltx" };

                // Create enumeration that only includes spreadsheet file extensions
                var enumeration = new FileSystemEnumerable<FileSystemInfo>(argument1,(ref FileSystemEntry entry) => entry.ToFileSystemInfo(),new EnumerationOptions() { RecurseSubdirectories = true })
                {
                    ShouldIncludePredicate = (ref FileSystemEntry entry) =>
                    {
                        
                        // Skip directories (is this necessary?)
                        if (entry.IsDirectory)
                        {
                            return false;
                        }
                        // End of skip directories

                        foreach (string extension in extensions)
                        {
                            var fileExtension = Path.GetExtension(entry.FileName);
                            if (fileExtension.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                            {
                                // Include the file if it matches extensions
                                return true;
                            }
                        }
                        // Doesn't match, exclude it
                        return false;
                    }
                };

                // Loop spreadsheets based on enumeration
                foreach (var file in enumeration.ToList())
                {

                    // Data types
                    int conv_dir_number = 1;
                    int conv_file_number = 1;
                    string conv_dir_sub = conv_dir + conv_dir_number;
                    string copy_new_filepath = conv_dir_sub + "\\" + "orgFile_" + file.Name;
                    string copy_new_filename = "orgFile_" + file.Name;
                    string conv_new_filepath = conv_dir_sub + "\\" + conv_file_number + ".xlsx";

                    // Create new subdirectory for the spreadsheet
                    while (Directory.Exists(@conv_dir_sub))
                    {
                        conv_dir_number++;
                        conv_dir_sub = conv_dir + conv_dir_number;
                    }
                    DirectoryInfo OutputDirSub = Directory.CreateDirectory(@conv_dir_sub);

                    // Copy spreadsheet
                    copy_new_filepath = conv_dir_sub + "\\" + "orgFile_" + file.Name;
                    File.Copy(file.FullName, copy_new_filepath);

                    try
                    {

                        // Convert spreadsheet
                        switch (file.Extension)
                        {

                            // OpenDocument file formats
                            case ".fods":
                            case ".ods":
                            case ".ots":
                                // Code to execute
                                numFAILED++;
                                success = false;

                                // Inform user
                                Console.WriteLine(file.FullName);
                                Console.WriteLine($"--> Conversion {success} - {error_message[3]}");

                                // Output result in open CSV file
                                var newLine2 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},,,,{success},{error_message[3]}");
                                csv.AppendLine(newLine2);
                                break;

                            // Legacy Microsoft Excel file formats
                            case ".xla":
                                // Code to execute
                                numFAILED++;
                                success = false;

                                // Inform user
                                Console.WriteLine(file.FullName);
                                Console.WriteLine($"--> Conversion {success} - {error_message[5]}");

                                // Output result in open CSV file
                                var newLine3 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},,,,{success},{error_message[5]}");
                                csv.AppendLine(newLine3);
                                break;
                            case ".xls":
                            case ".xlt":
                                // Code to execute
                                numFAILED++;
                                success = false;

                                // Inform user
                                Console.WriteLine(file.FullName);
                                Console.WriteLine($"--> Conversion {success} - {error_message[1]}");

                                // Output result in open CSV file
                                var newLine4 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},,,,{success},{error_message[1]}");
                                csv.AppendLine(newLine4);
                                break;

                                // Office Open XML file formats
                            case ".xlsb":
                                // Code to execute
                                numFAILED++;
                                success = false;

                                // Inform user
                                Console.WriteLine(file.FullName);
                                Console.WriteLine($"--> Conversion {success} - {error_message[2]}");

                                // Output result in open CSV file
                                var newLine5 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},,,,{success},{error_message[2]}");
                                csv.AppendLine(newLine5);
                                break;
                            case ".xlam":
                                // Code to execute
                                numFAILED++;
                                success = false;

                                // Inform user
                                Console.WriteLine(file.FullName);
                                Console.WriteLine($"--> Conversion {success} - {error_message[5]}");

                                // Output result in open CSV file
                                var newLine6 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},,,,{success},{error_message[5]}");
                                csv.AppendLine(newLine6);
                                break;
                            case ".xlsm":
                            case ".xltm":
                            case ".xltx":

                                // Loop naming of new conversion
                                while (File.Exists(conv_new_filepath))
                                {
                                    conv_file_number++;
                                    conv_new_filepath = conv_dir_sub + "\\" + conv_file_number + ".xlsx";
                                }
                                // This code must exist otherwise the switch will convert .fods to .xlsx by error
                                conv_new_filepath = conv_dir_sub + "\\" + conv_file_number + ".xlsx";

                                // Conversion code
                                byte[] byteArray = File.ReadAllBytes(copy_new_filepath);
                                using (MemoryStream stream = new MemoryStream())
                                {
                                    stream.Write(byteArray, 0, (int)byteArray.Length);
                                    using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                                    {
                                        spreadsheetDoc.ChangeDocumentType(DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
                                    }
                                    File.WriteAllBytes(conv_new_filepath, stream.ToArray());
                                }
                                success = true;

                                // Inform user
                                Console.WriteLine(file.FullName);
                                Console.WriteLine($"--> Conversion {success}");

                                // Output result in open CSV file
                                var newLine7 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},{conv_new_filepath},{conv_file_number}.xlsx,.xlsx,{success},{error_message[0]}");
                                csv.AppendLine(newLine7);
                                break;
                            case ".xlsx":
                                success = true;

                                // Copy spreadsheet
                                conv_new_filepath = conv_dir_sub + "\\" + conv_file_number + ".xlsx";
                                File.Copy(copy_new_filepath, conv_new_filepath);

                                // Inform user
                                Console.WriteLine(file.FullName);
                                Console.WriteLine($"--> Conversion {success} - {error_message[6]}");

                                // Output result in open CSV file
                                var newLine8 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},{conv_new_filepath},{conv_file_number}.xlsx,.xlsx,{success},{error_message[6]}");
                                csv.AppendLine(newLine8);
                                break;

                        }

                    }

                    catch (System.IO.FileFormatException)
                    {
                        // Code to execute
                        numFAILED++;
                        success = false;

                        // Inform user
                        Console.WriteLine(file.FullName);
                        Console.WriteLine($"--> Conversion {success} - {error_message[4]}");

                        // Output result in open CSV file
                        var newLine1 = string.Format($"{file.FullName},{file.Name},{file.Extension},{copy_new_filepath},{copy_new_filename},,,,{success},{error_message[4]}");
                        csv.AppendLine(newLine1);
                    }

                    finally
                    {

                    }

                }

            }
            else if (argument3 == "Recursive=No")
            {
                Console.WriteLine("Feature coming.");
            }
            else
            {
                Console.WriteLine("Invalid recursive argument");
            }

            // Close CSV file to log results
            string convert_CSV_filepath = results_directory + "\\2_Convert_Results.csv";
            File.WriteAllText(convert_CSV_filepath, csv.ToString());
            
            // Inform user of results
            int numCOMPLETE = numTOTAL - numFAILED;
            Console.WriteLine("---");
            Console.WriteLine($"{numCOMPLETE} out of {numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"{numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {convert_CSV_filepath}");
            Console.WriteLine("Conversion finished");
            Console.WriteLine("---");

        }

    }

}
