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
            string convert_directory = argument2 + "\\Converted_Spreadsheets";

            Console.WriteLine("Convert");

            // Create new folder
            if (Directory.Exists(@convert_directory))
            {
                Console.WriteLine($"Error: Directory identified: {convert_directory}");
                Console.WriteLine("ErrorMessage: Directory with name 'Converted_Spreadsheets' must not exist in the output directory to prevent accidental overwriting of data");
            }
            else
            {
                DirectoryInfo OutputDir = Directory.CreateDirectory(@argument2 + "\\Converted_Spreadsheets");

                // Copy spreadsheets
                if (argument3 == "Recursive=Yes")
                {
                    foreach (string dirPath in Directory.GetDirectories(@argument1, "*", SearchOption.AllDirectories))
                    {
                        //Copy all the files
                        foreach (string newPath in Directory.GetFiles(@argument1, "*.*", SearchOption.AllDirectories))
                            File.Copy(newPath, newPath.Replace(argument1, convert_directory));
                    }
                }
                else if (argument3 == "Recursive=No")
                {

                }
                else
                {
                    Console.WriteLine("Invalid argument in position args[3]");
                }
                // Rename
                // int filenumber = 1;
                // if (prefix has value)
                // {
                // filename = prefix + ++filenumber + ".xlsx"
                // }
                // else 
                // filename = ++filenumber + ".xlsx"

                // Convert spreadsheet
                //Console.WriteLine($"{} out of {numTOTAL} conversions completed");
                //var csv = new StringBuilder();
                //var newLine0 = string.Format($"#,{file_format[0]},{file_format_description[0]}");
                //csv.AppendLine(newLine0);
                //string convert_CSV_filename = argument2 + "\\2_Convert_Results.csv";
                //File.WriteAllText(convert_CSV_filename, csv.ToString());
                //Console.WriteLine($"Results saved to CSV log in filepath: {convert_CSV_filename}");
                Console.WriteLine("Conversion finished");
                Console.WriteLine("---");
            }

        }

    }

}
