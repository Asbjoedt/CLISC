using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{

    public partial class Spreadsheet
    {

        // Generate date to use in name of new directory
        public static string GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy.MM.dd");
        }
        public string dateStamp = GetTimestamp(DateTime.Now);

        // Create name for new results directory
        public string results_directory = "";
        public string Name_Directory(string argument1, string argument2)
        {
            // Identify available name for results directory
            int results_directory_number = 1;
            results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;
            while (Directory.Exists(@results_directory))
            {
                results_directory_number++;
                results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;
            }

            // Create results directory
            DirectoryInfo OutputDir = Directory.CreateDirectory(@results_directory);

            return results_directory;
        }

    }

}
