using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CLISC
{

    public partial class Spreadsheet
    {

        // Convert spreadsheets
        public void Convert(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("Convert");
            Console.WriteLine("---");

            // Open CSV file to log results
            string complete = "complete", fail = "fail";
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath,Original filename, Original file format,Conversion Complete,New filepath,New filename, New file format");
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
            string convert_directory = results_directory + "\\Spreadsheets\\";
            DirectoryInfo OutputDir = Directory.CreateDirectory(@convert_directory);

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
                foreach (var file in enumeration)
                {
                    //Create new subdirectory for each spreadsheet
                    int convert_directory_number = 1;
                    string convert_directory_sub = convert_directory + convert_directory_number;
                    while (Directory.Exists(@convert_directory_sub))
                    {
                        convert_directory_number++;
                        convert_directory_sub = convert_directory + convert_directory_number;
                    }
                    DirectoryInfo OutputDirSub = Directory.CreateDirectory(@convert_directory_sub);

                    // Rename new copy
                    int copy_file_number = 1;
                    string new_filepath = convert_directory_sub + "\\" + copy_file_number + file.Extension;
                    while (File.Exists(new_filepath))
                    {
                        copy_file_number++;
                        new_filepath = convert_directory_sub + "\\" + copy_file_number + file.Extension;
                    }

                    // Copy spreadsheet into subdirectory
                    File.Copy(file.FullName, new_filepath);

                    // Convert spreadsheet
                    switch (file.Extension)
                    {
                        
                        // OpenDocument file formats
                        case ".fods":
                        case ".ods":
                        case ".ots":
                            Console.WriteLine(file.FullName);
                            Console.WriteLine("- Error: Use LibreOffice to convert the spreadsheets");
                            break;

                        // Legacy Microsoft Excel file formats
                        case ".xla":
                        case ".xls":
                        case ".xlt":
                            Console.WriteLine(file.FullName);
                            Console.WriteLine("- Error: Legacy Excel spreadsheets cannot be converted. Feature on its way");
                            break;

                        // Office Open XML file formats
                        case ".xlsb":
                            Console.WriteLine(file.FullName);
                            Console.WriteLine("- Error: XLSB spreadsheets cannot be converted. Feature on its way");
                            break;
                        case ".xlam":
                        case ".xlsm":
                        case ".xlsx":
                        case ".xltm":
                        case ".xltx":
                            byte[] byteArray = File.ReadAllBytes(new_filepath);
                            using (MemoryStream stream = new MemoryStream())
                            {
                                stream.Write(byteArray, 0, (int)byteArray.Length);
                                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                                {
                                    spreadsheetDoc.ChangeDocumentType(DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
                                }
                                new_filepath = convert_directory_sub + "\\" + copy_file_number + ".xlsx";
                                File.WriteAllBytes(new_filepath, stream.ToArray());
                            }
                            break;

                    }

                    // Output result in open CSV file
                    var newLine1 = string.Format($"{file.FullName},{file.Name},{file.Extension},dd,{new_filepath},{copy_file_number}.xlsx,.xlsx");
                    csv.AppendLine(newLine1);
                }
            }
            else if (argument3 == "Recursive=No")
            {
                Console.WriteLine("ddd");
            }
            else
            {
                Console.WriteLine("Invalid recursive argument in position args[3]");
            }

            // Close CSV file to log results
            string convert_CSV_filepath = results_directory + "\\2_Convert_Results.csv";
            File.WriteAllText(convert_CSV_filepath, csv.ToString());
            //Console.WriteLine($"{} out of {numTOTAL} conversions completed");
            Console.WriteLine($"Results saved to CSV log in filepath: {convert_CSV_filepath}");
            Console.WriteLine("Conversion finished");
            Console.WriteLine("---");

        }

    }

}
