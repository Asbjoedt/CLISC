using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace CLISC
{

    public partial class Spreadsheet
    {

        // Convert spreadsheets
        public void Convert(string argument1, string argument2, string argument3)
        {
            string complete = "complete", fail = "fail";

            Console.WriteLine("Convert");

            // Create subdirectory
            int results_directory_number = 1;
            string results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
            while (Directory.Exists(@results_directory))
            {
                results_directory_number++;
                results_directory = argument2 + "\\CLISC_Results_" + results_directory_number;
            }
            string convert_directory = results_directory + "\\Converted_Spreadsheets";
            DirectoryInfo process = Directory.CreateDirectory(@convert_directory);

            // Copy spreadsheets to subdirectory
            if (argument3 == "Recursive=Yes")
            {

                var extensions = new[] { ".fods", ".ods", ".ots" };

                var files = (from file in Directory.EnumerateFiles(@argument1)
                             where extensions.Contains(Path.GetExtension(file), StringComparer.InvariantCultureIgnoreCase)
                             select new
                             {
                                 Source = file,
                                 Destination = Path.Combine(@convert_directory, Path.GetFileName(file))
                             });

                // old stuff
                foreach (var file in files)
                {
                    File.Copy(file.Source, file.Destination);
                }



                foreach (string dirPath in Directory.GetDirectories(@argument1, "*", SearchOption.AllDirectories))
                {
                    //Copy all the files
                    foreach (string newPath in Directory.GetFiles(@argument1, "*.*", SearchOption.AllDirectories))
                        File.Copy(newPath, newPath.Replace(argument1, results_directory));

                    // Rename
                    // int filenumber = 1;

                    // Output results in CSV
                    // var csv = new StringBuilder();
                    // var newLine0 = string.Format($"Original filepath,Original filename, New filepath,New filename, Conversion Completed");
                    // csv.AppendLine(newLine0);

                    // loop nednestående newLine1 = 1++
                    // var newLine1 = string.Format($"{},{},{convert_directory},{}");
                    // csv.AppendLine(newLine1);
                    // string count_CSV_filepath = results_directory + "\\2_Convert_Results.csv";
                    // File.WriteAllText(convert_CSV_filepath, csv.ToString());
                    // Console.WriteLine($"Results saved to CSV log in filepath: {convert_CSV_filepath}");
                }
            }
            else if (argument3 == "Recursive=No")
            {

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
