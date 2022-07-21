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

        // Conversion error messages
        int numCOMPLETE = 0;
        int numFAILED = 0;
        bool convert_success;
        string[] convert_error_message = { "", "Legacy Excel file formats are not supported", "Binary XLSB file format is not supported", "LibreOffice is not installed in filepath: C:\\Program Files\\LibreOffice", "Spreadsheet is password protected or corrupt", "Microsoft Excel Add-In file format is not supported", "Spreadsheet is already .xlsx file format. File was copied and renamed" };

        // Convert spreadsheets method
        public void Convert(string argument1, string argument2, string argument3, string argument4)
        {

            Console.WriteLine("CONVERT");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original file format;New copy filepath;New copy filename; New convert filepath; New convert filename; New convert file format;Success;Message");
            csv.AppendLine(newLine0);

            // Create subdirectory for spreadsheets
            string conv_dir = results_directory + "\\docCollection\\";
            DirectoryInfo OutputDir = Directory.CreateDirectory(@conv_dir);

            // Create enumeration of original spreadsheets based on input directory
            IEnumerable<string> org_enumeration = Enumerate_Original<string>(argument1, argument3);

            // Loop spreadsheets based on enumeration
            foreach (var file in org_enumeration.ToList()) // Is .ToList() necessary?
            {

                // Create instance for finding file information
                FileInfo file_info = new FileInfo(file);

                // Create data types for original spreadsheets
                string org_extension = file_info.Extension;
                string org_filename = file_info.Name;
                string org_filepath = file_info.FullName;

                // Create data types for converted spreadsheets
                int conv_dir_number = 1;
                int conv_file_number = 1;
                string conv_dir_sub = conv_dir + conv_dir_number;
                string conv_new_filename = conv_file_number + ".xlsx";
                string conv_new_filepath = conv_dir_sub + "\\" + conv_new_filename;

                // Create data types for copied original spreadsheets
                string copy_new_filename = "orgFile_" + org_filename;
                string copy_new_filepath = conv_dir_sub + "\\" + copy_new_filename;

                // Create new subdirectory for the spreadsheet
                while (Directory.Exists(@conv_dir_sub))
                {
                    conv_dir_number++;
                    conv_dir_sub = conv_dir + conv_dir_number;
                }
                DirectoryInfo OutputDirSub = Directory.CreateDirectory(@conv_dir_sub);

                // Copy spreadsheet
                copy_new_filepath = conv_dir_sub + "\\" + "orgFile_" + org_filename;
                File.Copy(org_filepath, copy_new_filepath);

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
                            convert_success = Convert_OpenDocument(org_filepath, copy_new_filepath, conv_dir_sub);

                            if (convert_success == false)
                            {
                                numFAILED++;

                                // Output result in open CSV file
                                var newLine2 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};;;;{convert_success};{convert_error_message[3]}");
                                csv.AppendLine(newLine2);
                            }

                            // The next line must exist otherwise CSV will have wrong "conv_new_filepath"
                            conv_new_filepath = conv_dir_sub + "\\" + conv_file_number + ".xlsx";

                            // Output result in open CSV file
                            var newLine9 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};{conv_new_filepath};{conv_file_number}.xlsx;.xlsx;{convert_success};{convert_error_message[0]}");
                            csv.AppendLine(newLine9);

                            break;

                        // Legacy Microsoft Excel file formats
                        case ".xla":

                            // Conversion code
                            numFAILED++;
                            convert_success = false;

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {convert_error_message[5]}");

                            // Output result in open CSV file
                            var newLine3 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};;;;{convert_success};{convert_error_message[5]}");
                            csv.AppendLine(newLine3);

                            break;

                        case ".xls":
                        case ".xlt":

                            // Conversion code
                            numFAILED++;
                            convert_success = false;

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {convert_error_message[1]}");

                            // Output result in open CSV file
                            var newLine4 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};;;;{convert_success};{convert_error_message[1]}");
                            csv.AppendLine(newLine4);

                            break;

                        // Office Open XML file formats
                        case ".xlsb":

                            // Conversion code
                            numFAILED++;
                            convert_success = false;

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {convert_error_message[2]}");

                            // Output result in open CSV file
                            var newLine5 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};;;;{convert_success};{convert_error_message[2]}");
                            csv.AppendLine(newLine5);

                            break;

                        case ".xlam":

                            // Conversion code
                            numFAILED++;
                            convert_success = false;

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {convert_error_message[5]}");

                            // Output result in open CSV file
                            var newLine6 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};;;;{convert_success};{convert_error_message[5]}");
                            csv.AppendLine(newLine6);

                            break;

                        case ".xlsm":
                        case ".xltm":
                        case ".xltx":

                            // The next line must exist otherwise the switch will not convert
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

                            convert_success = true;

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success}");

                            // Output result in open CSV file
                            var newLine7 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};{conv_new_filepath};{conv_file_number}.xlsx;.xlsx;{convert_success};{convert_error_message[0]}");
                            csv.AppendLine(newLine7);

                            break;

                        case ".xlsx":

                            // No converison
                            convert_success = true;

                            // Copy spreadsheet
                            conv_new_filepath = conv_dir_sub + "\\" + conv_file_number + ".xlsx";
                            File.Copy(copy_new_filepath, conv_new_filepath);

                            // Inform user
                            Console.WriteLine(org_filepath);
                            Console.WriteLine($"--> Conversion {convert_success} - {convert_error_message[6]}");

                            // Output result in open CSV file
                            var newLine8 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};{conv_new_filepath};{conv_file_number}.xlsx;.xlsx;{convert_success};{convert_error_message[6]}");
                            csv.AppendLine(newLine8);

                            break;

                    }

                }

                catch (FileFormatException)
                {
                    // Code to execute
                    numFAILED++;
                    convert_success = false;

                    // Inform user
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Conversion {convert_success} - {convert_error_message[4]}");

                    // Output result in open CSV file
                    var newLine1 = string.Format($"{org_filepath};{org_filename};{org_extension};{copy_new_filepath};{copy_new_filename};;;;{convert_success};{convert_error_message[4]}");
                    csv.AppendLine(newLine1);
                }

                finally
                {

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
