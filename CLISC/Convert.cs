using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using System.IO.Enumeration;

namespace CLISC
{

    public partial class Spreadsheet
    {

        // Convert spreadsheets
        public void Convert(string argument1, string argument2, string argument3)
        {

            Console.WriteLine("Convert");

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
            string convert_directory = results_directory + "\\Converted_Spreadsheets\\";
            DirectoryInfo process = Directory.CreateDirectory(@convert_directory);

            // Copy spreadsheets to subdirectory


            if (argument3 == "Recursive=Yes")
            {
                var extensions = new List<string> { ".fods", ".ods", ".ots", ".xls", ".xlt", ".xlam", ".xlsb", ".xlsm", ".xlsx", ".xltm", ".xltx" };

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
                    // Rename new copy
                    int copy_file_number = 1;
                    string new_filepath = convert_directory + copy_file_number + file.Extension;
                    while (File.Exists(new_filepath))
                    {
                        copy_file_number++;
                        new_filepath = convert_directory + copy_file_number + file.Extension;
                    }

                    // Copy
                    File.Copy(file.FullName, new_filepath);


                    //string complete = "complete", fail = "fail";
                }

                // Rename
                // int filenumber = 1;

                // Output results in CSV
                // var csv = new StringBuilder();
                // var newLine0 = string.Format($"Original filepath,Original filename, Original filesize, New filepath,New filename, New filesize, Conversion Completed");
                // csv.AppendLine(newLine0);

                // loop nednestående newLine1 = 1++
                // var newLine1 = string.Format($"{},{},{convert_directory},{}");
                // csv.AppendLine(newLine1);
                // string count_CSV_filepath = results_directory + "\\2_Convert_Results.csv";
                // File.WriteAllText(convert_CSV_filepath, csv.ToString());
                // Console.WriteLine($"Results saved to CSV log in filepath: {convert_CSV_filepath}");
            }
            else if (argument3 == "Recursive=No")
            {
                Console.WriteLine("ddd");
            }
            else
            {
                Console.WriteLine("Invalid recursive argument in position args[3]");
            }


            //Console.WriteLine($"{} out of {numTOTAL} conversions completed");
            //Console.WriteLine($"Results saved to CSV log in filepath: {convert_CSV_filename}");
            Console.WriteLine("Conversion finished");
            Console.WriteLine("---");

        }

    }

}
