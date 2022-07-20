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
        public void Archive(string argument1, string argument2)
        {
            // Open CSV file to log results
            var csv = new StringBuilder();
            var newLine0 = string.Format($"Original filepath;Original filesize (KB);Original checksum;Conversion identified;Conversion filepath;Conversion filesize (KB);Conversion checksum");
            csv.AppendLine(newLine0);

            // Validate file format standards
            switch (file.Extension)
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
                    Validate_OOXML(argument1, argument2);
                    break;
            }

            // Zip the output directory
            ZIP_Directory(argument1, argument2);

            // Close CSV file to log results
            string archive_CSV_filepath = results_directory + "\\4_Archive_Results.csv";
            File.WriteAllText(archive_CSV_filepath, csv.ToString());

            // Inform user of results
            Console.WriteLine("---");
            Console.WriteLine("X spreadsheets failed file format validation");
            Console.WriteLine($"x out of {numTOTAL} spreadsheets were archived");
            Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Archiving finished");
            Console.WriteLine("---");

        }

    }

}
