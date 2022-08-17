using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    public partial class Compare
    {
        public bool Compare_Workbook(string Results_Directory, string file_folder, string org_filepath, string conv_filepath)
        {
            //Create "Beyond Compare 4" script file
            string bcscript_filepath = Results_Directory + "\\bcscript.txt";
            string bcscript_results_filepath = file_folder + "\\comparisonResults.txt";
            using (StreamWriter bcscript = File.CreateText(bcscript_filepath))
            {
                    bcscript.WriteLine($"data-report layout:interleaved options:display-mismatches output-to:\"{bcscript_results_filepath}\" \"{org_filepath}\" \"{conv_filepath}\"");
            }

            // Use Beyond Compare 4 command line for comparison
            Process app = new Process();
            app.StartInfo.FileName = "C:\\Program Files\\Beyond Compare 4\\BCompare.exe";
            app.StartInfo.Arguments = $"\"@{bcscript_filepath}\" /silent";
            app.Start();
            app.WaitForExit();
            app.Close();

            // Read logfile to identify differences
            var compare_message = File.ReadAllText(bcscript_results_filepath);
            Console.WriteLine(compare_message);

            // Delete BC4 results file
            if (File.Exists(bcscript_results_filepath))
            {
                File.Delete(bcscript_results_filepath);
            }
            // Delete BC script
            if (File.Exists(bcscript_filepath))
            {
                File.Delete(bcscript_filepath);
            }

            bool compare_success = true;

            //If there is workbook differences
            if (compare_success == false)
            {
                numTOTAL_diff++;
            }
            return compare_success;
        }
    }
}
