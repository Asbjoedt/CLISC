using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Public data type needed for Compare.cs
        public string bcscript_filepath = "";

        public string Compare_Workbook(string argument0, string Results_Directory, string docCollection, string org_filepath, string conv_filepath)
        {
            //Create "Beyond Compare" script file
            bcscript_filepath = Results_Directory + "\\bcscript.txt";
            string bcscript_results_filepath = docCollection + "\\comparisonResults.txt";
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
            compare_message = File.ReadAllText(bcscript_results_filepath);

            // Delete logfile
            if (File.Exists(bcscript_results_filepath))
            {
                File.Delete(bcscript_results_filepath);
            }

            // Delete BC script
            if (File.Exists(bcscript_filepath))
            {
                File.Delete(bcscript_filepath);
            }

            bool compare_success;

            // If there is workbook differences
            //if (comparison_success == fail)
            //{
            //    numTOTAL_diff++;
            //
            //    // Inform user
            //    Console.WriteLine(compare_conv_filepath);
            //    Console.WriteLine($"--> Comparison {success} - Workbook differences identified");
            //}

            // No workbook differences
            //else
            //{
            //    // Inform user
            //    Console.WriteLine($"--> Comparison {success}");
            //}

            // If archiving, create supplementary comparison logfile in the docCollection subdir
            if (argument0 == "Count&Convert&Compare&Archive")
            {
                bcscript_results_filepath = docCollection + "\\comparisonResults.html";
                using (StreamWriter bcscript = File.CreateText(bcscript_filepath))
                {
                    bcscript_results_filepath = docCollection + "\\comparisonResults.html";
                    bcscript.WriteLine($"data-report layout:interleaved options:display-mismatches title:CLISC_Comparison_Results output-to:\"{bcscript_results_filepath}\" output-options:wrap-word,html-color \"{org_filepath}\" \"{conv_filepath}\"");
                }
                // Use Beyond Compare 4 command line for comparison
                app = new Process();
                app.StartInfo.FileName = "C:\\Program Files\\Beyond Compare 4\\BCompare.exe";
                app.StartInfo.Arguments = $"\"@{bcscript_filepath}\" /silent";
                app.Start();
                app.WaitForExit();
                app.Close();
                // Delete BC script
                if (File.Exists(bcscript_filepath))
                {
                    File.Delete(bcscript_filepath);
                }
            }

            return compare_message;
        }
    }
}
