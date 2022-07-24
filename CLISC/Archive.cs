using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Archive the spreadsheets according to advanced archival requirements
        public void Archive(string argument0, string argument1, string argument2, string Results_Directory, List<fileIndex> File_List)
        {
            Console.WriteLine("ARCHIVE");
            Console.WriteLine("---");

            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filename;Original checksum;Copy filepath; Copy filename; Conversion identified;Conversion filepath;Conversion filename;Conversion checksum;File format validated, Data quality message");
            csv.AppendLine(newLine0);

            // Create enumeration of files in each folder of docCollection_enumeration
            foreach (var folder in File_List)
            {
                var subdir_enumeration = new List<string>();
                subdir_enumeration = (List<string>)Directory.EnumerateFiles(folder, "*", SearchOption.TopDirectoryOnly)
                    .ToList();
                foreach (var file in subdir_enumeration)
                {

                    // Rename and move converted spreadsheets

                    // Copy original spreadsheets


                    // Validate file format standards
                    switch (this._Conv_Extension)
                    {

                        // Validate OpenDocument file formats
                        case ".fods":
                        case ".ods":
                        case ".ots":
                            break;

                        // Validate Office Open XML file formats
                        case ".xlam":
                        case ".xlsm":
                        case ".xlsx":
                        case ".xltx":
                            validation_message = Validate_OOXML(conv_filepath);
                            break;

                            default:
                            Console.WriteLine(file_info.FullName);
                            Console.WriteLine("--> The file format is not supported in validation routine");
                            break;
                    }

                    // Calculate checksums
                    string copy_checksum = Calculate_MD5(org_filepath);
                    string conv_checksum = Calculate_MD5(conv_filepath);

                    // Check for data quality requirements
                    string dataquality_message = "";

                    // Output result in open CSV file
                    //var newLine1 = string.Format($"{org_filepath};{org_filename};{copy_filepath};{copy_filename};{convert_success};{conv_filepath};{conv_filename};{conv_checksum};{validation_message};{dataquality_message}");
                    //csv.AppendLine(newLine1);
                }
            }

            // Close CSV file to log results
            string CSV_filepath = Results_Directory + "\\4_Archive_Results.csv";
            File.WriteAllText(CSV_filepath, csv.ToString());

            // Zip the output directory
            ZIP_Directory(Results_Directory);

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine("X spreadsheets failed file format validation");
            Console.WriteLine($"x out of {numTOTAL} spreadsheets were archived");
            Console.WriteLine($"Results saved to CSV log in filepath: {CSV_filepath}"); //Filepath is incorrect. It is not the zipped path
            Console.WriteLine("Archiving finished");
            Console.WriteLine("---");
        }
    }
}
