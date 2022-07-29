using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Create public data types
        public string Results_Directory = "";
        public string CSV_filepath = "";

        // Generate date to use in name of new directory
        public static string GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy.MM.dd");
        }
        public string dateStamp = GetTimestamp(DateTime.Now);
        
        // Create name for new results directory
        public string Create_Directory_Results(string outputdir)
        {
            // Identify available name for results directory
            int results_directory_number = 1;
            Results_Directory = outputdir + "\\CLISC_" + dateStamp;
            while (Directory.Exists(Results_Directory))
            {
                results_directory_number++;
                Results_Directory = outputdir + "\\CLISC_" + dateStamp + "_" + results_directory_number;
            }
            // Create results directory
            DirectoryInfo OutputDir = Directory.CreateDirectory(Results_Directory);

            return Results_Directory;
        }
    }
}
